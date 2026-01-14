# -*- coding: utf-8 -*-

import sys
import os
import ctypes
import tempfile
import time
import threading
import json
import logging
import webbrowser
import base64
import winsound
import re
from urllib import request, parse, error
from functools import wraps
import wx
import addonHandler
import globalPluginHandler
import config
import gui
import ui
import api
import scriptHandler
import tones

# --- Import Welcome Wizard ---
try:
    from . import welcome
except ImportError:
    welcome = None

# --- 1. Path Setup ---
lib_dir = os.path.join(os.path.dirname(__file__), "lib")
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# --- 2. Library Import ---
sr_error = None
try:
    import speech_recognition as sr
except Exception as e:
    sr = None
    sr_error = str(e)
    logging.error(f"TypeTalkAI: Library Import Error: {e}")

addonHandler.initTranslation()

# --- 3. Security & Data ---

_SECURE_DATA = {
    "gpt_5_1": "aHR0cHM6Ly9ocmlkb3ktbXVsdGktbW9kZWwtYXBpLnZlcmNlbC5hcHAvP21vZGVsPUdQVC01LTE=",
    "gpt_5_1_chat": "aHR0cHM6Ly9ocmlkb3ktbXVsdGktbW9kZWwtYXBpLnZlcmNlbC5hcHAvP21vZGVsPUdQVC01LTEtQ2hhdA==",
    "gemini_3_0": "aHR0cHM6Ly9ocmlkb3ktbXVsdGktbW9kZWwtYXBpLnZlcmNlbC5hcHAvP21vZGVsPUdvb2dsZS1HZW1pbmktMy0wLVBybw=="
}

def _get_secure_url(key):
    try:
        encoded = _SECURE_DATA.get(key, "")
        return base64.b64decode(encoded).decode("utf-8")
    except:
        return ""

ALL_LANGUAGES = [
    ("Afrikaans", "af-ZA"), ("Albanian", "sq-AL"), ("Amharic", "am-ET"), ("Arabic (Saudi Arabia)", "ar-SA"),
    ("Armenian", "hy-AM"), ("Azerbaijani", "az-AZ"), ("Bengali (Bangladesh)", "bn-BD"), ("Bengali (India)", "bn-IN"),
    ("Bulgarian", "bg-BG"), ("Catalan", "ca-ES"), ("Chinese (Mandarin)", "zh-CN"), ("Croatian", "hr-HR"),
    ("Czech", "cs-CZ"), ("Danish", "da-DK"), ("Dutch", "nl-NL"), ("English (Australia)", "en-AU"),
    ("English (India)", "en-IN"), ("English (UK)", "en-GB"), ("English (US)", "en-US"), ("Filipino", "fil-PH"),
    ("Finnish", "fi-FI"), ("French", "fr-FR"), ("German", "de-DE"), ("Greek", "el-GR"), ("Gujarati", "gu-IN"),
    ("Hebrew", "he-IL"), ("Hindi", "hi-IN"), ("Hungarian", "hu-HU"), ("Icelandic", "is-IS"), ("Indonesian", "id-ID"),
    ("Italian", "it-IT"), ("Japanese", "ja-JP"), ("Javanese", "jv-ID"), ("Kannada", "kn-IN"), ("Khmer", "km-KH"),
    ("Korean", "ko-KR"), ("Latin", "la"), ("Latvian", "lv-LV"), ("Malay", "ms-MY"), ("Malayalam", "ml-IN"),
    ("Marathi", "mr-IN"), ("Myanmar (Burmese)", "my-MM"), ("Nepali", "ne-NP"), ("Norwegian", "no-NO"),
    ("Persian", "fa-IR"), ("Polish", "pl-PL"), ("Portuguese", "pt-PT"), ("Punjabi", "pa-IN"), ("Romanian", "ro-RO"),
    ("Russian", "ru-RU"), ("Serbian", "sr-RS"), ("Sinhala", "si-LK"), ("Slovak", "sk-SK"), ("Spanish", "es-ES"),
    ("Sundanese", "su-ID"), ("Swahili", "sw-KE"), ("Swedish", "sv-SE"), ("Tamil", "ta-IN"), ("Telugu", "te-IN"),
    ("Thai", "th-TH"), ("Turkish", "tr-TR"), ("Ukrainian", "uk-UA"), ("Urdu", "ur-PK"), ("Vietnamese", "vi-VN")
]

PUNCTUATION_MAP = {
    r"\bদাড়ি\b": "।", r"\bকমা\b": ",", r"\bপ্রশ্নবোধক চিহ্ন\b": "?", r"\bআশ্চর্যবোধক চিহ্ন\b": "!",
    r"\bকোলন\b": ":", r"\bসেমিকোলন\b": ";", r"\bনতুন লাইন\b": "\n",
    r"\bcomma\b": ",", r"\bfull stop\b": ".", r"\bperiod\b": ".", r"\bquestion mark\b": "?",
    r"\bexclamation mark\b": "!", r"\bcolon\b": ":", r"\bsemicolon\b": ";", r"\bnew line\b": "\n",
    r"\bपूर्ण विराम\b": "।", r"\bअल्पविराम\b": ",", r"\bप्रश्नवाचक\b": "?"
}

HELP_TEXT = """HOW TO USE TYPETALK AI:

1. First, Activate Command Mode by pressing:
   NVDA + Shift + Space

2. Then, press one of the following keys:

   S  : Start / Stop Dictation
   A  : Toggle AI Processing
   P  : Toggle Spoken Punctuation (Offline)
   T  : Toggle Translation
   E  : Toggle Emoji Support
   I  : Change Input Language
   L  : Change Translation Target
   M  : Change AI Model
   D  : About Developer
   H  : Show this Help Menu

Note: Spoken Punctuation only works when AI is OFF.
"""

LINKS = {
    "Hridoy Modding Hub": "https://t.me/Hridoy_Modding_Hub",
    "School of mind light": "https://t.me/schoolofmindlight2018",
    "Omnisent community": "https://t.me/omnisent25",
    "Helpful app store": "https://t.me/Helpfulappstore"
}

ABOUT_TITLE = "TypeTalk AI v1.0"
ABOUT_SUB = "Powered by AI • Developed by Md Hridoy Sheikh"
ABOUT_BODY = (
    "TypeTalk AI is an intelligent, real-time voice typing assistant built specifically for NVDA users. "
    "It goes far beyond basic dictation by leveraging advanced artificial intelligence to convert speech "
    "into clear, accurate, and well-structured text instantly.\n\n"
    "With features such as grammar correction, automatic punctuation, seamless translation, and expressive "
    "emoji support, TypeTalk AI helps users communicate more effectively and confidently.\n\n"
    "Developer:\nMd Hridoy Sheikh\nDhaka, Bangladesh\n\n"
    "Mission:\nTo utilize programming and technology to improve the quality of life for visually impaired individuals."
)

confspec = {
    "input_language": "string(default='bn-BD')",
    "use_ai_processing": "boolean(default=True)",
    "use_spoken_punctuation": "boolean(default=False)",
    "ai_model": "string(default='gpt_5_1')",
    "use_translation": "boolean(default=False)",
    "target_language": "string(default='English (US)')",
    "use_emoji": "boolean(default=False)",
    "custom_url": "string(default='')",
    "custom_method": "string(default='GET')",
    "custom_param": "string(default='q')",
    "custom_response_path": "string(default='response')",
    "last_welcome_version": "string(default='0.0.0')"
}

AI_MODELS = [
    ("GPT 5.1 (Fast)", "gpt_5_1"),
    ("GPT 5.1 Chat", "gpt_5_1_chat"),
    ("Google Gemini 3.0 Pro", "gemini_3_0"),
    ("Custom API...", "custom")
]

def play_sound(sound_name):
    file_path = os.path.join(os.path.dirname(__file__), f"{sound_name}.wav")
    if os.path.exists(file_path):
        winsound.PlaySound(file_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        return
    if sound_name == "start": tones.beep(1000, 50) 
    elif sound_name == "stop": tones.beep(500, 80)
    elif sound_name == "success": tones.beep(1500, 100)
    elif sound_name == "error": tones.beep(150, 400)

class AboutDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="About TypeTalk AI", size=(600, 550))
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        title = wx.StaticText(self, label=ABOUT_TITLE)
        title.SetFont(wx.Font(16, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        mainSizer.Add(title, 0, wx.ALIGN_CENTER | wx.TOP, 15)
        sub = wx.StaticText(self, label=ABOUT_SUB)
        sub.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_ITALIC, wx.FONTWEIGHT_NORMAL))
        mainSizer.Add(sub, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        mainSizer.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.ALL, 5)
        desc = wx.TextCtrl(self, value=ABOUT_BODY, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_CENTER)
        mainSizer.Add(desc, 1, wx.EXPAND | wx.ALL, 15)
        gridSizer = wx.GridSizer(rows=2, cols=2, vgap=5, hgap=5)
        for name, link in LINKS.items():
            btn = wx.Button(self, label=name)
            btn.Bind(wx.EVT_BUTTON, lambda evt, url=link: self.openLink(url))
            gridSizer.Add(btn, 0, wx.EXPAND)
        mainSizer.Add(gridSizer, 0, wx.EXPAND | wx.ALL, 10)
        closeBtn = wx.Button(self, wx.ID_OK, label="Close")
        mainSizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 15)
        self.SetSizer(mainSizer)
        self.Center()
    def openLink(self, url):
        try: webbrowser.open(url)
        except: wx.MessageBox(f"Link: {url}", "Info")

class HelpDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="TypeTalk AI Help", size=(500, 500))
        sizer = wx.BoxSizer(wx.VERTICAL)
        text = wx.TextCtrl(self, value=HELP_TEXT, style=wx.TE_MULTILINE | wx.TE_READONLY)
        text.SetFont(wx.Font(12, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        sizer.Add(text, 1, wx.EXPAND | wx.ALL, 10)
        closeBtn = wx.Button(self, wx.ID_OK, label="Close")
        sizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.SetSizer(sizer)
        self.Center()

class UpdateDialog(wx.Dialog):
    def __init__(self, parent, version, changelog, download_url):
        super().__init__(parent, title="Critical Update", size=(400, 300), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP | wx.CENTER)
        self.download_url = download_url
        self.version = version
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        warn = wx.StaticText(self, label=f"Update Required ({version})")
        warn.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        warn.SetForegroundColour(wx.Colour(200, 0, 0))
        sizer.Add(warn, 0, wx.ALIGN_CENTER | wx.TOP, 15)
        
        log = wx.TextCtrl(self, value=changelog, style=wx.TE_MULTILINE | wx.TE_READONLY)
        sizer.Add(log, 1, wx.EXPAND | wx.ALL, 10)
        
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        updateBtn = wx.Button(self, label="Update Now")
        updateBtn.Bind(wx.EVT_BUTTON, self.onUpdate)
        updateBtn.SetDefault()
        
        closeBtn = wx.Button(self, wx.ID_CANCEL, label="Later (Locked)")
        
        btnSizer.Add(updateBtn, 0, wx.RIGHT, 10)
        btnSizer.Add(closeBtn, 0)
        
        sizer.Add(btnSizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 15)
        self.SetSizer(sizer)
        self.CenterOnScreen()

    def onUpdate(self, evt):
        self.EndModal(wx.ID_OK)
        self.progressDlg = wx.ProgressDialog("Downloading Update", "Connecting to server...", maximum=100, parent=gui.mainFrame, style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        t = threading.Thread(target=self._download_worker)
        t.daemon = True
        t.start()

    def _download_worker(self):
        try:
            temp_dir = tempfile.gettempdir()
            file_name = f"TypeTalkAI_{self.version}.nvda-addon"
            file_path = os.path.join(temp_dir, file_name)
            
            with request.urlopen(self.download_url) as response:
                total_size = int(response.info().get('Content-Length', 0))
                downloaded = 0
                block_size = 4096 
                
                with open(file_path, 'wb') as out_file:
                    while True:
                        buffer = response.read(block_size)
                        if not buffer: break
                        downloaded += len(buffer)
                        out_file.write(buffer)
                        
                        if total_size > 0:
                            percent = int((downloaded / total_size) * 100)
                            wx.CallAfter(self.progressDlg.Update, percent, f"Downloading: {percent}%")
            
            wx.CallAfter(self.progressDlg.Destroy)
            wx.CallAfter(os.startfile, file_path)
            
        except Exception as e:
            wx.CallAfter(self.progressDlg.Destroy)
            wx.CallAfter(wx.MessageBox, f"Download Failed: {str(e)}", "Error", wx.ICON_ERROR)

class TypeTalkSettingsPanel(gui.settingsDialogs.SettingsPanel):
    title = "TypeTalk AI"
    def makeSettings(self, settingsSizer):
        sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        
        voiceGroup = wx.StaticBox(self, label="Voice Input Settings")
        voiceSizer = wx.StaticBoxSizer(voiceGroup, wx.VERTICAL)
        vHelper = gui.guiHelper.BoxSizerHelper(voiceGroup, sizer=voiceSizer)
        lang_names = [x[0] for x in ALL_LANGUAGES]
        self.inputLang = vHelper.addLabeledControl("Input Language:", wx.Choice, choices=lang_names)
        try:
            current = config.conf["TypeTalkAI"]["input_language"]
            idx = next(i for i, v in enumerate(ALL_LANGUAGES) if v[1] == current)
            self.inputLang.SetSelection(idx)
        except: self.inputLang.SetSelection(0)
        sHelper.addItem(voiceSizer)

        aiGroup = wx.StaticBox(self, label="AI & Processing Logic")
        aiSizer = wx.StaticBoxSizer(aiGroup, wx.VERTICAL)
        aHelper = gui.guiHelper.BoxSizerHelper(aiGroup, sizer=aiSizer)
        self.useAI = aHelper.addItem(wx.CheckBox(aiGroup, label="Enable AI Detection"))
        self.useAI.Value = config.conf["TypeTalkAI"]["use_ai_processing"]
        self.useAI.Bind(wx.EVT_CHECKBOX, self.onToggleAI)
        self.usePunctuation = aHelper.addItem(wx.CheckBox(aiGroup, label="Enable Spoken Punctuation (Offline)"))
        self.usePunctuation.Value = config.conf["TypeTalkAI"]["use_spoken_punctuation"]
        self.usePunctuation.Bind(wx.EVT_CHECKBOX, self.onTogglePunctuation)
        model_choices = [x[0] for x in AI_MODELS]
        self.aiModel = aHelper.addLabeledControl("Select AI Model:", wx.Choice, choices=model_choices)
        try:
            current = config.conf["TypeTalkAI"]["ai_model"]
            idx = next(i for i, v in enumerate(AI_MODELS) if v[1] == current)
            self.aiModel.SetSelection(idx)
        except: self.aiModel.SetSelection(0)
        self.aiModel.Bind(wx.EVT_CHOICE, self.onModelChange)
        transSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.useTrans = wx.CheckBox(aiGroup, label="Translate to:")
        self.useTrans.Value = config.conf["TypeTalkAI"]["use_translation"]
        transSizer.Add(self.useTrans, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.targetLang = wx.Choice(aiGroup, choices=lang_names)
        try: 
            cur_trans = config.conf["TypeTalkAI"]["target_language"]
            idx = lang_names.index(cur_trans)
            self.targetLang.SetSelection(idx)
        except: self.targetLang.SetSelection(0)
        transSizer.Add(self.targetLang, 1, wx.EXPAND)
        aHelper.addItem(transSizer)
        self.useEmoji = aHelper.addItem(wx.CheckBox(aiGroup, label="Add Emojis"))
        self.useEmoji.Value = config.conf["TypeTalkAI"]["use_emoji"]
        sHelper.addItem(aiSizer)

        customGroup = wx.StaticBox(self, label="Custom API Configuration")
        customSizer = wx.StaticBoxSizer(customGroup, wx.VERTICAL)
        cHelper = gui.guiHelper.BoxSizerHelper(customGroup, sizer=customSizer)
        self.custUrl = cHelper.addLabeledControl("API URL:", wx.TextCtrl)
        self.custUrl.Value = config.conf["TypeTalkAI"]["custom_url"]
        methodSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.custMethod = wx.Choice(customGroup, choices=["GET", "POST"])
        self.custMethod.SetSelection(0 if config.conf["TypeTalkAI"]["custom_method"] == "GET" else 1)
        methodSizer.Add(self.custMethod, 0, wx.RIGHT, 15)
        self.custParam = wx.TextCtrl(customGroup)
        self.custParam.Value = config.conf["TypeTalkAI"]["custom_param"]
        methodSizer.Add(self.custParam, 1, wx.EXPAND)
        cHelper.addItem(methodSizer)
        self.custPath = cHelper.addLabeledControl("JSON Path:", wx.TextCtrl)
        self.custPath.Value = config.conf["TypeTalkAI"]["custom_response_path"]
        sHelper.addItem(customSizer)
        self.aboutBtn = sHelper.addItem(wx.Button(self, label="About Developer"))
        self.aboutBtn.Bind(wx.EVT_BUTTON, self.onAbout)
        
        self.onToggleAI(None)
        self.onModelChange(None)

    def onToggleAI(self, evt):
        en = self.useAI.Value
        if en:
            self.usePunctuation.SetValue(False)
            self.usePunctuation.Enable(False)
        else:
            self.usePunctuation.Enable(True)
        self.aiModel.Enable(en)
        self.useTrans.Enable(en)
        self.targetLang.Enable(en)
        self.useEmoji.Enable(en)
        if not en: self.enableCustomFields(False)
        else: self.onModelChange(None)

    def onTogglePunctuation(self, evt):
        if self.usePunctuation.Value:
            if self.useAI.Value:
                wx.MessageBox("Please disable AI Detection first.", "Conflict")
                self.usePunctuation.SetValue(False)

    def onModelChange(self, evt):
        if not self.useAI.Value: return
        sel_idx = self.aiModel.GetSelection()
        code = AI_MODELS[sel_idx][1]
        self.enableCustomFields(code == "custom")

    def enableCustomFields(self, enable):
        self.custUrl.Enable(enable)
        self.custMethod.Enable(enable)
        self.custParam.Enable(enable)
        self.custPath.Enable(enable)

    def onAbout(self, evt):
        dlg = AboutDialog(self)
        dlg.ShowModal()
        dlg.Destroy()

    def onSave(self):
        config.conf["TypeTalkAI"]["input_language"] = ALL_LANGUAGES[self.inputLang.GetSelection()][1]
        config.conf["TypeTalkAI"]["use_ai_processing"] = self.useAI.Value
        config.conf["TypeTalkAI"]["use_spoken_punctuation"] = self.usePunctuation.Value
        config.conf["TypeTalkAI"]["ai_model"] = AI_MODELS[self.aiModel.GetSelection()][1]
        config.conf["TypeTalkAI"]["use_translation"] = self.useTrans.Value
        config.conf["TypeTalkAI"]["target_language"] = ALL_LANGUAGES[self.targetLang.GetSelection()][0]
        config.conf["TypeTalkAI"]["use_emoji"] = self.useEmoji.Value
        config.conf["TypeTalkAI"]["custom_url"] = self.custUrl.Value
        config.conf["TypeTalkAI"]["custom_method"] = "GET" if self.custMethod.GetSelection() == 0 else "POST"
        config.conf["TypeTalkAI"]["custom_param"] = self.custParam.Value
        config.conf["TypeTalkAI"]["custom_response_path"] = self.custPath.Value

def finally_(func, final):
    @wraps(func)
    def new(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        finally:
            final()
    return new

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = "TypeTalk AI"
    
    is_recording = False
    toggling = False 
    temp_audio_file = os.path.join(tempfile.gettempdir(), "typetalk_audio.wav")
    
    update_available = False
    update_locked = False 
    REPO_API = "https://api.github.com/repos/hridoyhub/TypeTalk-AI-NVDA-Addon/releases/latest"
    
    latest_ver = ""
    latest_log = ""
    latest_url = ""

    def __init__(self):
        super(GlobalPlugin, self).__init__()
        config.conf.spec["TypeTalkAI"] = confspec
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(TypeTalkSettingsPanel)
        
        # --- DYNAMIC VERSION CHECK (FIXED) ---
        try:
            # Get version directly from manifest
            current_addon_version = addonHandler.getCodeAddon().manifest['version']
        except:
            current_addon_version = "1.0.0"

        last_seen = config.conf["TypeTalkAI"]["last_welcome_version"]
        
        if last_seen != current_addon_version and welcome:
            # Pass current version to save it later
            wx.CallAfter(self._show_welcome_wizard, current_addon_version)
        
        threading.Thread(target=self._check_for_update, daemon=True).start()

    def _show_welcome_wizard(self, version_to_save):
        try:
            welcome.show_wizard()
            # Save the DYNAMIC version to config
            config.conf["TypeTalkAI"]["last_welcome_version"] = version_to_save
        except: pass

    def _check_for_update(self):
        try:
            cur_ver = addonHandler.getCodeAddon().manifest['version']
            req = request.Request(self.REPO_API, headers={'User-Agent': 'NVDA-Addon'})
            with request.urlopen(req, timeout=10) as response:
                data = json.loads(response.read().decode('utf-8'))
                
            remote_tag = data.get('tag_name', 'v0.0').lstrip('v')
            changelog = data.get('body', 'No details.')
            assets = data.get('assets', [])
            
            if remote_tag != cur_ver and remote_tag > cur_ver:
                download_url = ""
                for asset in assets:
                    if asset['name'].endswith('.nvda-addon'):
                        download_url = asset['browser_download_url']
                        break
                
                if download_url:
                    self.update_available = True
                    self.update_locked = True
                    self.latest_ver = remote_tag
                    self.latest_log = changelog
                    self.latest_url = download_url
                    wx.CallAfter(self._show_update_dialog)

        except Exception as e:
            logging.error(f"TypeTalkAI Update Check Failed: {e}")

    def _show_update_dialog(self):
        if self.latest_url:
            dlg = UpdateDialog(gui.mainFrame, self.latest_ver, self.latest_log, self.latest_url)
            dlg.ShowModal()
            dlg.Destroy()

    def terminate(self):
        try: gui.settingsDialogs.NVDASettingsDialog.categoryClasses.remove(TypeTalkSettingsPanel)
        except: pass

    def _show_error(self, spoken_msg, log_msg):
        play_sound("error")
        ui.message(spoken_msg)
        logging.error(f"TypeTalkAI Error: {log_msg}")
        
    def getScript(self, gesture):
        if not self.toggling:
            return super(GlobalPlugin, self).getScript(gesture)
        
        script = super(GlobalPlugin, self).getScript(gesture)
        if not script:
            return finally_(self.script_invalidCommand, self.finish)
            
        return finally_(script, self.finish)

    def finish(self):
        self.toggling = False
        self.clearGestureBindings()
        self.bindGestures(self.__gestures)

    def script_error(self, gesture):
        play_sound("error")

    def script_invalidCommand(self, gesture):
        ui.message("Command unavailable. Press H for Help.")
        
    def script_updateLocked(self, gesture):
        play_sound("error")
        ui.message("Update Required! Please update to continue.")
        if self.latest_url:
            wx.CallAfter(self._show_update_dialog)

    @scriptHandler.script(description="Activates Command Layer")
    def script_activateLayer(self, gesture):
        if self.update_locked:
            self.script_updateLocked(gesture)
            return

        if self.toggling:
            self.script_error(gesture)
            return
        self.bindGestures(self.__VisionGestures)
        self.toggling = True
        play_sound("start")

    @scriptHandler.script(description="Toggle Voice Typing")
    def script_smartDictation(self, gesture):
        if self.toggling: self.finish()
        if not sr:
            self._show_error("Library missing", f"SpeechRecognition not found. {sr_error}")
            return

        try:
            if not self.is_recording:
                self.is_recording = True
                tones.beep(800, 100) 
                
                ctypes.windll.winmm.mciSendStringW('close all', None, 0, 0)
                res = ctypes.windll.winmm.mciSendStringW('open new type waveaudio alias myaudio', None, 0, 0)
                if res != 0: raise Exception("Mic Init Failed")
                ctypes.windll.winmm.mciSendStringW('record myaudio', None, 0, 0)
                
                lang_code = config.conf["TypeTalkAI"]["input_language"]
                lang_name = next((name for name, code in ALL_LANGUAGES if code == lang_code), "Unknown")
                ui.message(f"Listening ({lang_name})...")
            else:
                self.is_recording = False
                tones.beep(400, 100)
                
                ctypes.windll.winmm.mciSendStringW(f'save myaudio "{self.temp_audio_file}"', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('close myaudio', None, 0, 0)
                ui.message("Processing...")
                threading.Thread(target=self._process_pipeline, daemon=True).start()
        except Exception as e:
            self.is_recording = False
            self._show_error("Mic Error", str(e))

    @scriptHandler.script(description="Toggle AI Processing")
    def script_toggleAI(self, gesture):
        if self.toggling: self.finish()
        current = config.conf["TypeTalkAI"]["use_ai_processing"]
        if not current: config.conf["TypeTalkAI"]["use_spoken_punctuation"] = False
        new_val = not current
        config.conf["TypeTalkAI"]["use_ai_processing"] = new_val
        ui.message(f"AI Detection {'Enabled' if new_val else 'Disabled'}")

    @scriptHandler.script(description="Toggle Spoken Punctuation (Offline)")
    def script_togglePunctuation(self, gesture):
        if self.toggling: self.finish()
        if config.conf["TypeTalkAI"]["use_ai_processing"]:
            ui.message("Disable AI first")
            return
        current = config.conf["TypeTalkAI"]["use_spoken_punctuation"]
        new_val = not current
        config.conf["TypeTalkAI"]["use_spoken_punctuation"] = new_val
        ui.message(f"Spoken Punctuation {'Enabled' if new_val else 'Disabled'}")

    @scriptHandler.script(description="Toggle Translation")
    def script_toggleTranslation(self, gesture):
        if self.toggling: self.finish()
        if not config.conf["TypeTalkAI"]["use_ai_processing"]:
            ui.message("Enable AI first")
            return
        current = config.conf["TypeTalkAI"]["use_translation"]
        new_val = not current
        config.conf["TypeTalkAI"]["use_translation"] = new_val
        ui.message(f"Translation {'Enabled' if new_val else 'Disabled'}")

    @scriptHandler.script(description="Toggle Emoji")
    def script_toggleEmoji(self, gesture):
        if self.toggling: self.finish()
        if not config.conf["TypeTalkAI"]["use_ai_processing"]:
            ui.message("Enable AI first")
            return
        current = config.conf["TypeTalkAI"]["use_emoji"]
        new_val = not current
        config.conf["TypeTalkAI"]["use_emoji"] = new_val
        ui.message(f"Emoji {'Enabled' if new_val else 'Disabled'}")

    @scriptHandler.script(description="Show Developer Info")
    def script_showAbout(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_about_dialog)
    
    @scriptHandler.script(description="Show Help Menu")
    def script_showHelp(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_help_dialog)

    def _show_about_dialog(self):
        try:
            dlg = AboutDialog(gui.mainFrame)
            dlg.ShowModal()
            dlg.Destroy()
        except: pass

    def _show_help_dialog(self):
        try:
            dlg = HelpDialog(gui.mainFrame)
            dlg.ShowModal()
            dlg.Destroy()
        except: pass

    @scriptHandler.script(description="Change AI Model")
    def script_changeModel(self, gesture):
        if self.toggling: self.finish()
        if not config.conf["TypeTalkAI"]["use_ai_processing"]:
            ui.message("Enable AI first")
            return
        wx.CallAfter(self._show_model_dialog)

    def _show_model_dialog(self):
        try:
            choices = [x[0] for x in AI_MODELS]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, "Select AI Model:", "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"]["ai_model"]
            try:
                idx = next(i for i, v in enumerate(AI_MODELS) if v[1] == current)
                dlg.SetSelection(idx)
            except: pass
            dlg.Center()
            dlg.Raise()
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection()
                code = AI_MODELS[sel_idx][1]
                name = AI_MODELS[sel_idx][0]
                config.conf["TypeTalkAI"]["ai_model"] = code
                wx.CallLater(100, ui.message, f"Model: {name}")
            dlg.Destroy()
        except: pass

    @scriptHandler.script(description="Change Input Language")
    def script_changeInputLang(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_input_dialog)
        
    def _show_input_dialog(self):
        try:
            choices = [x[0] for x in ALL_LANGUAGES]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, "Select Input Language:", "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"]["input_language"]
            try:
                idx = next(i for i, v in enumerate(ALL_LANGUAGES) if v[1] == current)
                dlg.SetSelection(idx)
            except: pass
            dlg.Center()
            dlg.Raise()
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection()
                code = ALL_LANGUAGES[sel_idx][1]
                name = ALL_LANGUAGES[sel_idx][0]
                config.conf["TypeTalkAI"]["input_language"] = code
                wx.CallLater(100, ui.message, f"Input: {name}")
            dlg.Destroy()
        except: pass

    @scriptHandler.script(description="Change Translation Language")
    def script_changeTransLang(self, gesture):
        if self.toggling: self.finish()
        if not config.conf["TypeTalkAI"]["use_ai_processing"]:
            ui.message("Enable AI first")
            return
        wx.CallAfter(self._show_trans_dialog)

    def _show_trans_dialog(self):
        try:
            choices = [x[0] for x in ALL_LANGUAGES]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, "Select Translation Target:", "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"]["target_language"]
            try:
                name_list = [x[0] for x in ALL_LANGUAGES]
                idx = name_list.index(current)
                dlg.SetSelection(idx)
            except: pass
            dlg.Center()
            dlg.Raise()
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection()
                name = ALL_LANGUAGES[sel_idx][0]
                config.conf["TypeTalkAI"]["target_language"] = name
                wx.CallLater(100, ui.message, f"Target: {name}")
            dlg.Destroy()
        except: pass

    def _process_pipeline(self):
        try:
            if not os.path.exists(self.temp_audio_file): return
            r = sr.Recognizer()
            with sr.AudioFile(self.temp_audio_file) as source:
                audio_data = r.record(source)
            lang_code = config.conf["TypeTalkAI"]["input_language"]
            try:
                raw_text = r.recognize_google(audio_data, language=lang_code)
            except sr.UnknownValueError:
                play_sound("error")
                wx.CallAfter(ui.message, "No speech detected")
                return
            except Exception as e:
                wx.CallAfter(self._show_error, "Network Error", str(e))
                return

            if not config.conf["TypeTalkAI"]["use_ai_processing"]:
                if config.conf["TypeTalkAI"]["use_spoken_punctuation"]:
                    processed_text = self._apply_spoken_punctuation(raw_text)
                    wx.CallAfter(self._paste_text, processed_text)
                else:
                    formatted_text = self._offline_format(raw_text)
                    wx.CallAfter(self._paste_text, formatted_text)
            else:
                wx.CallAfter(ui.message, "Refining...")
                self._call_ai_api(raw_text)
        except Exception as e:
            wx.CallAfter(self._show_error, "System Error", str(e))
        finally:
            try: os.remove(self.temp_audio_file)
            except: pass

    def _apply_spoken_punctuation(self, text):
        for pattern, replacement in PUNCTUATION_MAP.items():
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
        text = text.strip()
        if text: text = text[0].upper() + text[1:]
        return text

    def _offline_format(self, text):
        text = text.strip()
        if not text: return ""
        text = text[0].upper() + text[1:]
        if text[-1] not in ".!?|": text += "."
        return text

    def _call_ai_api(self, input_text):
        try:
            model_code = config.conf["TypeTalkAI"]["ai_model"]
            use_trans = config.conf["TypeTalkAI"]["use_translation"]
            target_lang = config.conf["TypeTalkAI"]["target_language"]
            use_emoji = config.conf["TypeTalkAI"]["use_emoji"]
            prompt = ""
            if use_trans:
                prompt = f"Translate to {target_lang}: '{input_text}'. Output ONLY translated text."
            else:
                prompt = f"Fix grammar: '{input_text}'. Output ONLY fixed text."
            if use_emoji: prompt += " Add emojis."

            api_url = ""
            method = "GET"
            param_name = "q"
            resp_path = "response"

            if model_code == "custom":
                api_url = config.conf["TypeTalkAI"]["custom_url"]
                method = config.conf["TypeTalkAI"]["custom_method"]
                param_name = config.conf["TypeTalkAI"]["custom_param"]
                resp_path = config.conf["TypeTalkAI"]["custom_response_path"]
            else:
                api_url = _get_secure_url(model_code)
                method = "GET"
                param_name = "q"
                resp_path = "response"

            if not api_url: raise Exception("Invalid Model or Missing URL")
            final_text = ""
            if method == "GET":
                query_string = parse.urlencode({param_name: prompt})
                full_url = f"{api_url}&{query_string}" if "?" in api_url else f"{api_url}?{query_string}"
                req = request.Request(full_url, headers={'User-Agent': 'Mozilla/5.0'})
                with request.urlopen(req, timeout=15) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    final_text = data.get(resp_path, "Error: Field not found")
            else:
                data_json = json.dumps({param_name: prompt}).encode('utf-8')
                req = request.Request(api_url, data=data_json, method="POST", headers={'Content-Type': 'application/json'})
                with request.urlopen(req, timeout=15) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    final_text = data.get(resp_path, "Error: Field not found")
            if final_text: wx.CallAfter(self._paste_text, final_text)
            else: wx.CallAfter(self._show_error, "AI Error", "Empty response")
        except Exception as e:
            wx.CallAfter(self._show_error, "API Error", str(e))

    def _paste_text(self, text):
        play_sound("success")
        api.copyToClip(text)
        try:
            user32 = ctypes.windll.user32
            VK_CONTROL = 0x11
            VK_V = 0x56
            KEYEVENTF_KEYUP = 0x0002
            user32.keybd_event(VK_CONTROL, 0, 0, 0)
            user32.keybd_event(VK_V, 0, 0, 0)
            user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        except: pass
        wx.CallLater(500, ui.message, f"Typed: {text}")

    __gestures = {
        "kb:NVDA+shift+space": "activateLayer",
    }
    
    __VisionGestures = {
        "kb:s": "smartDictation",
        "kb:a": "toggleAI",
        "kb:t": "toggleTranslation",
        "kb:e": "toggleEmoji",
        "kb:p": "togglePunctuation",
        "kb:d": "showAbout",
        "kb:h": "showHelp",
        "kb:m": "changeModel",
        "kb:i": "changeInputLang",
        "kb:l": "changeTransLang",
    }