# -*- coding: utf-8 -*-
import ctypes
import os
import wx

import globalPluginHandler
import ui
import config
import addonHandler
from gui.settingsDialogs import SettingsPanel as _SettingsPanel
from gui import guiHelper
from nvwave import playWaveFile

# Log at INFO so it shows at default log level
try:
	from logHandler import log
except Exception:
	class _Dummy:
		def info(self, *a, **k): pass
	log = _Dummy()

addonHandler.initTranslation()

ALLOWED_STEPS = (1, 5, 10, 15, 20)

# ---------- Config spec ----------
# Keep section name 'batteryMon' for compatibility with existing configs
if 'batteryMon' not in config.conf.spec:
	config.conf.spec['batteryMon'] = {
		# Separate step sizes for each power state
		'stepSizeOnBattery': 'integer(default=5, min=1, max=20)',
		'stepSizeOnAC': 'integer(default=5, min=1, max=20)',
		# Poll cadence in seconds (hidden setting; 10–3600)
		'pollSeconds': 'integer(default=15, min=10, max=3600)',
		# Output options
		'speak': 'boolean(default=True)',
		'playSound': 'boolean(default=False)',
		# Separate sound paths for down/up (no built-in fallbacks)
		'soundPathDown': 'string(default="")',
		'soundPathUp': 'string(default="")',
		# Back-compat (ignored if the two above are set)
		'stepSize': 'integer(default=5, min=1, max=20)',
	}

def _normalizeStep(value, fallback=5):
	try:
		v = int(value)
	except Exception:
		return fallback
	return v if v in ALLOWED_STEPS else fallback

def _stepDown():
	conf = config.conf['batteryMon']
	v = conf.get('stepSizeOnBattery', None)
	if v is None:
		v = conf.get('stepSize', 5)
	return _normalizeStep(v, 5)

def _stepUp():
	conf = config.conf['batteryMon']
	v = conf.get('stepSizeOnAC', None)
	if v is None:
		v = conf.get('stepSize', 5)
	return _normalizeStep(v, 5)

# ---------- Windows power status ----------
class SYSTEM_POWER_STATUS(ctypes.Structure):
	_fields_ = [
		("ACLineStatus", ctypes.c_ubyte),
		("BatteryFlag", ctypes.c_ubyte),
		("BatteryLifePercent", ctypes.c_ubyte),
		("SystemStatusFlag", ctypes.c_ubyte),
		("BatteryLifeTime", ctypes.c_uint32),
		("BatteryFullLifeTime", ctypes.c_uint32),
	]

def _getPower():
	s = SYSTEM_POWER_STATUS()
	if not ctypes.windll.kernel32.GetSystemPowerStatus(ctypes.byref(s)):
		return (False, None)
	onAC = (s.ACLineStatus == 1)
	percent = None if s.BatteryLifePercent == 255 else int(s.BatteryLifePercent)
	return (onAC, percent)

# ---------- Settings panel ----------
class BatteryMonSettingsPanel(_SettingsPanel):
	title = _("Battery Utils")

	def makeSettings(self, sizer):
		self.helper = guiHelper.BoxSizerHelper(self, sizer=sizer)

		# Battery (discharging) step
		self.stepBatt = self.helper.addLabeledControl(
			_("Announce every percent decrease, on power"),
			wx.Choice,
			choices=[str(x) for x in ALLOWED_STEPS]
		)
		try:
			self.stepBatt.SetStringSelection(str(_stepDown()))
		except Exception:
			self.stepBatt.SetSelection(1)  # default to "5"

		# AC (charging) step
		self.stepAC = self.helper.addLabeledControl(
			_("Announce every percent increase, when plugged in"),
			wx.Choice,
			choices=[str(x) for x in ALLOWED_STEPS]
		)
		try:
			self.stepAC.SetStringSelection(str(_stepUp()))
		except Exception:
			self.stepAC.SetSelection(1)

		# Toggles
		self.speakChk = self.helper.addItem(wx.CheckBox(self, label=_("Speak alerts")))
		self.speakChk.SetValue(bool(config.conf['batteryMon']['speak']))

		self.soundChk = self.helper.addItem(wx.CheckBox(self, label=_("Play sounds for alerts")))
		self.soundChk.SetValue(bool(config.conf['batteryMon']['playSound']))

		# Decrease sound (down)
		rowDown = wx.BoxSizer(wx.HORIZONTAL)
		lblDown = wx.StaticText(self, label=_("Sound file for decreased battery, must be in wav format"))
		rowDown.Add(lblDown, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

		self.soundPathDownTxt = wx.TextCtrl(self, value=config.conf['batteryMon']['soundPathDown'], size=(300, -1))
		rowDown.Add(self.soundPathDownTxt, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

		self.browseDownBtn = wx.Button(self, label=_("Browse…"))
		self.browseDownBtn.Bind(wx.EVT_BUTTON, self._onBrowseDown)
		rowDown.Add(self.browseDownBtn, 0, wx.ALIGN_CENTER_VERTICAL)

		self.helper.sizer.Add(rowDown, 0, wx.EXPAND | wx.TOP, 8)

		# Increase sound (up)
		rowUp = wx.BoxSizer(wx.HORIZONTAL)
		lblUp = wx.StaticText(self, label=_("Sound file for increased battery, must be in wav format"))
		rowUp.Add(lblUp, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

		self.soundPathUpTxt = wx.TextCtrl(self, value=config.conf['batteryMon']['soundPathUp'], size=(300, -1))
		rowUp.Add(self.soundPathUpTxt, 1, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 6)

		self.browseUpBtn = wx.Button(self, label=_("Browse…"))
		self.browseUpBtn.Bind(wx.EVT_BUTTON, self._onBrowseUp)
		rowUp.Add(self.browseUpBtn, 0, wx.ALIGN_CENTER_VERTICAL)

		self.helper.sizer.Add(rowUp, 0, wx.EXPAND | wx.TOP, 8)

	def _onBrowseGeneric(self, textCtrl: wx.TextCtrl):
		curPath = textCtrl.GetValue().strip()
		# Default to the current file's directory if valid, else user's home
		defaultDir = None
		if curPath:
			try:
				cp = os.path.dirname(curPath)
				if cp and os.path.isdir(cp):
					defaultDir = cp
			except Exception:
				pass
		if not defaultDir:
			defaultDir = os.path.expanduser("~")

		with wx.FileDialog(
			self,
			_("Choose sound file"),
			defaultDir=defaultDir,
			wildcard=_("Wave files (*.wav)|*.wav|All files (*.*)|*.*"),
			style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
		) as dlg:
			if dlg.ShowModal() == wx.ID_OK:
				textCtrl.SetValue(dlg.GetPath())

	def _onBrowseDown(self, evt):
		self._onBrowseGeneric(self.soundPathDownTxt)

	def _onBrowseUp(self, evt):
		self._onBrowseGeneric(self.soundPathUpTxt)

	def onSave(self):
		conf = config.conf['batteryMon']
		try:
			conf['stepSizeOnBattery'] = int(self.stepBatt.GetStringSelection())
		except Exception:
			conf['stepSizeOnBattery'] = 5
		try:
			conf['stepSizeOnAC'] = int(self.stepAC.GetStringSelection())
		except Exception:
			conf['stepSizeOnAC'] = 5
		conf['speak'] = bool(self.speakChk.GetValue())
		conf['playSound'] = bool(self.soundChk.GetValue())
		conf['soundPathDown'] = self.soundPathDownTxt.GetValue().strip()
		conf['soundPathUp'] = self.soundPathUpTxt.GetValue().strip()

# Register settings panel
import gui as _gui
_gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(BatteryMonSettingsPanel)

# ---------- Global plugin ----------
class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	scriptCategory = _("Battery Utils")

	def __init__(self, *a, **k):
		super(GlobalPlugin, self).__init__(*a, **k)
		self._timer = wx.PyTimer(self._tick)
		self._lastOnAC = None
		# Separate targets for down (discharging) and up (charging)
		self._nextDownTarget = None  # e.g., 40, 35, 30… or 40,39,38… if 1%
		self._nextUpTarget = None    # e.g., 45, 50… or 44,45,46… if 1%
		self._start()
		log.info("batteryMon: plugin initialized")

	def terminate(self):
		try:
			self._timer.Stop()
		except Exception:
			pass
		return super(GlobalPlugin, self).terminate()

	def _start(self):
		secs = max(10, min(3600, int(config.conf['batteryMon']['pollSeconds'])))
		self._timer.Start(secs * 1000)

	def _floorTo(self, p, step):
		return max(0, (p // step) * step)

	def _ceilTo(self, p, step):
		return min(100, ((p + step - 1) // step) * step)

	def _armDown(self, percent):
		step = _stepDown()
		self._nextDownTarget = self._floorTo(percent, step)
		log.info(f"batteryMon: armed (discharge), next target {self._nextDownTarget}% (step {step})")

	def _armUp(self, percent):
		step = _stepUp()
		self._nextUpTarget = self._ceilTo(percent, step)
		log.info(f"batteryMon: armed (charge), next target {self._nextUpTarget}% (step {step})")

	def _play(self, direction: str):
		"""
		direction: 'down' or 'up'
		Only plays user-chosen absolute paths; no built-in fallbacks.
		"""
		conf = config.conf['batteryMon']
		if not conf.get('playSound', False):
			return

		path = (conf.get('soundPathDown', '') if direction == 'down' else conf.get('soundPathUp', '')).strip()
		if not path:
			return
		# Normalize non-absolute or ~ paths; require .wav and existing file
		path = os.path.abspath(os.path.expanduser(path))
		if not path.lower().endswith(".wav") or not os.path.isfile(path):
			return

		try:
			playWaveFile(path)
		except Exception:
			pass

	def _sayDown(self, p):
		conf = config.conf['batteryMon']
		if conf.get('speak', True):
			ui.message(_("Alert, your battery has decreased to {p} percent.").format(p=p))
		self._play('down')
		log.info(f"batteryMon: announced (down) {p}%")

	def _sayUp(self, p):
		conf = config.conf['batteryMon']
		if conf.get('speak', True):
			ui.message(_("Alert, your battery has increased to {p} percent.").format(p=p))
		self._play('up')
		log.info(f"batteryMon: announced (up) {p}%")

	def _tick(self):
		onAC, percent = _getPower()
		if percent is None:
			return

		# AC state change?
		if self._lastOnAC is None or onAC != self._lastOnAC:
			self._lastOnAC = onAC
			if onAC:
				self._armUp(percent)
				self._nextDownTarget = None
			else:
				self._armDown(percent)
				self._nextUpTarget = None

		if onAC:
			step = _stepUp()
			while self._nextUpTarget is not None and percent >= self._nextUpTarget:
				self._sayUp(self._nextUpTarget)
				nextT = self._nextUpTarget + step
				self._nextUpTarget = nextT if nextT <= 100 else None
		else:
			step = _stepDown()
			while self._nextDownTarget is not None and percent <= self._nextDownTarget:
				self._sayDown(self._nextDownTarget)
				nextT = self._nextDownTarget - step
				self._nextDownTarget = nextT if nextT >= 0 else None

	# ---- Test gesture: reports next alert depending on state ----
	def script_testAnnounce(self, gesture):
		onAC, percent = _getPower()
		if percent is None:
			return
		if onAC:
			step = _stepUp()
			target = self._nextUpTarget if self._nextUpTarget is not None else self._ceilTo(percent, step)
			ui.message(_("Charging. Next alert at {t}% (now at {p}%).").format(t=target, p=percent))
		else:
			step = _stepDown()
			target = self._nextDownTarget if self._nextDownTarget is not None else self._floorTo(percent, step)
			ui.message(_("On battery. Next alert at {t}% (now at {p}%).").format(t=target, p=percent))

	__gestures = {
		# NVDA key is Caps Lock or Insert
		"kb:NVDA+alt+shift+b": "testAnnounce"
	}
