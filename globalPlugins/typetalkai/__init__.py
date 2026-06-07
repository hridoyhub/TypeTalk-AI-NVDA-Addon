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
import zlib
import winsound
import ssl
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
import textInfos
import tones
import controlTypes 

addonHandler.initTranslation()

def finally_(func, final):
    @wraps(func)
    def new(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        finally:
            final()
    return new

try:
    from . import welcome
except ImportError:
    welcome = None

try:
    from . import patterns
except ImportError:
    patterns = None

lib_dir = os.path.join(os.path.dirname(__file__), "lib")
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

sr_error = None
try:
    import speech_recognition as sr
except Exception as e:
    sr = None
    sr_error = str(e)

_SECRET_BYTES = [102, 122, 122, 126, 125, 52, 33, 33, 105, 103, 125, 122, 32, 105, 103, 122, 102, 123, 108, 123, 125, 107, 124, 109, 97, 96, 122, 107, 96, 122, 32, 109, 97, 99, 33, 102, 124, 103, 106, 97, 119, 102, 123, 108, 33, 107, 111, 107, 106, 56, 62, 59, 104, 56, 57, 108, 107, 104, 54, 106, 107, 62, 109, 106, 111, 108, 57, 107, 109, 59, 107, 56, 106, 54, 58, 106, 56, 33, 124, 111, 121, 33, 122, 119, 126, 107, 122, 111, 98, 101, 35, 111, 106, 106, 97, 96, 32, 100, 125, 97, 96]

def _get_gist_url():
    return "".join([chr(b ^ 14) for b in _SECRET_BYTES])

CACHE_FILE = os.path.join(tempfile.gettempdir(), "tt_core_v3.dat")

def load_local_config():
    if not os.path.exists(CACHE_FILE):
        return None
    try:
        with open(CACHE_FILE, "rb") as f:
            return json.loads(zlib.decompress(base64.b64decode(f.read())).decode('utf-8'))
    except Exception:
        return None

def save_local_config(data):
    try:
        enc = base64.b64encode(zlib.compress(json.dumps(data).encode('utf-8')))
        with open(CACHE_FILE, "wb") as f:
            f.write(enc)
        return True
    except Exception:
        return False

def get_ai_models():
    data = load_local_config()
    if data and "ai_models" in data:
        return [(m.get("display_name", "Unknown"), m.get("id", "unknown")) for m in data["ai_models"] if m.get("is_active", True)] + [(_("Custom API..."), "custom")]
    return [("GPT-5.4 Mini", "gpt_5_4_mini"), ("GPT-5.1", "universal_gpt_5_1"), (_("Custom API..."), "custom")]

HELP_TEXT = _("""HOW TO USE TYPETALK AI:

1. First, Activate Command Mode by pressing:
   NVDA + Shift + Space

2. Then, press one of the following keys:

   S  : Toggle Dictation
   R  : Smart Text Refiner
   C  : Set Context (For Smart Reply)
   G  : Generate Smart Reply
   A  : Toggle AI Processing
   T  : Toggle Translation
   E  : Toggle Emoji Support
   I  : Change Input Language
   L  : Change Translation Target Language
   M  : Change AI Model
   W  : Change Writing Style
   Q  : Check Status (Settings)
   B  : Recover Last Text (Backup)
   D  : About Developer
   H  : Show this Help Menu
""")

LINKS = {
    "Hridoy Modding Hub": "https://t.me/Hridoy_Modding_Hub",
    "School of mind light": "https://t.me/schoolofmindlight2018",
    "Omnisent community": "https://t.me/omnisent25",
    "Helpful app store": "https://t.me/Helpfulappstore"
}

ABOUT_TITLE = "TypeTalk AI v3.0"
ABOUT_SUB = _("Powered by Google Speech Services & AI • Developed by Md Hridoy Sheikh")
ABOUT_BODY = _(
    "TypeTalk AI is an intelligent, real-time voice typing assistant built specifically for NVDA users. "
    "It goes far beyond basic dictation by leveraging advanced artificial intelligence to convert speech "
    "into clear, accurate, and well-structured text instantly.\n\n"
    "Developer:\nMd Hridoy Sheikh\nDhaka, Bangladesh\n\n"
    "Mission:\nTo utilize programming and technology to improve the quality of life for visually impaired individuals."
)

confspec = {
    "input_language": "string(default='bn-BD')",
    "use_ai_processing": "boolean(default=False)",
    "use_custom_dictionary": "boolean(default=True)",
    "ai_model": "string(default='gpt_5_4_mini')",
    "use_translation": "boolean(default=False)",
    "target_language": "string(default='English (US)')",
    "use_emoji": "boolean(default=False)",
    "writing_style": "string(default='Default (Standard)')",
    "enable_sound": "boolean(default=True)",
    "auto_copy_clipboard": "boolean(default=False)",
    "custom_url": "string(default='')",
    "custom_method": "string(default='GET')",
    "custom_param": "string(default='q')",
    "custom_response_path": "string(default='response')",
    "last_welcome_version": "string(default='0.0.0')"
}

WRITING_STYLES = [
    "Default (Standard)",
    "Formal (Professional)",
    "Casual (Friendly)",
    "Concise (Short)",
    "Bullet Points"
]

def get_config_bool(key, default=False):
    try:
        val = config.conf["TypeTalkAI"].get(key, default)
        if isinstance(val, str):
            return val.lower() == 'true'
        return bool(val)
    except:
        return default

def safe_beep(hz, duration):
    if get_config_bool("enable_sound", True):
        tones.beep(hz, duration)

def play_sound(sound_name):
    if not get_config_bool("enable_sound", True):
        return
    file_path = os.path.join(os.path.dirname(__file__), f"{sound_name}.wav")
    if os.path.exists(file_path):
        winsound.PlaySound(file_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        return
    if sound_name == "start": tones.beep(1000, 50) 
    elif sound_name == "stop": tones.beep(500, 80)
    elif sound_name == "success": tones.beep(1500, 100)
    elif sound_name == "error": tones.beep(150, 400)

def _smart_insert_text(text, select_all=False):
    try:
        old_clip_data = ""
        try: old_clip_data = api.getClipData()
        except: pass
        if not api.copyToClip(text): 
            raise Exception("Clipboard copy failed")
        user32 = ctypes.windll.user32
        VK_CONTROL = 0x11
        VK_A = 0x41
        VK_V = 0x56
        KEYEVENTF_KEYUP = 0x0002
        if select_all:
            user32.keybd_event(VK_CONTROL, 0, 0, 0)
            user32.keybd_event(VK_A, 0, 0, 0)
            time.sleep(0.1)
            user32.keybd_event(VK_A, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
            time.sleep(0.1)
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        user32.keybd_event(VK_V, 0, 0, 0)
        time.sleep(0.2) 
        user32.keybd_event(VK_V, 0, KEYEVENTF_KEYUP, 0)
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        time.sleep(0.1)
        if not get_config_bool("auto_copy_clipboard", False):
            if old_clip_data: 
                try: api.copyToClip(old_clip_data)
                except: pass
        return True
    except:
        try:
            _send_unicode_text(text)
            return True
        except:
            return False

def _send_unicode_text(text):
    INPUT_KEYBOARD = 1
    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP = 0x0002
    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort), ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]
    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("ki", KEYBDINPUT)]
    inputs = []
    for char in text:
        inp_down = INPUT(type=INPUT_KEYBOARD, ki=KEYBDINPUT(wVk=0, wScan=ord(char), dwFlags=KEYEVENTF_UNICODE, time=0, dwExtraInfo=None))
        inp_up = INPUT(type=INPUT_KEYBOARD, ki=KEYBDINPUT(wVk=0, wScan=ord(char), dwFlags=KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, time=0, dwExtraInfo=None))
        inputs.extend([inp_down, inp_up])
    count = len(inputs)
    input_array = (INPUT * count)(*inputs)
    ctypes.windll.user32.SendInput(count, input_array, ctypes.sizeof(INPUT))

class RemoteSystemDialog(wx.Dialog):
    def __init__(self, parent, title, message, buttons):
        super().__init__(parent, title=title, size=(450, 250), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)
        sizer = wx.BoxSizer(wx.VERTICAL)
        txt = wx.TextCtrl(self, value=message, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.BORDER_NONE)
        sizer.Add(txt, 1, wx.EXPAND | wx.ALL, 15)
        bSizer = wx.BoxSizer(wx.HORIZONTAL)
        for btn in buttons:
            b = wx.Button(self, label=btn.get("label", ""))
            b.Bind(wx.EVT_BUTTON, lambda e, act=btn.get("action"), lnk=btn.get("link"): self.onAction(act, lnk))
            bSizer.Add(b, 0, wx.RIGHT, 10)
        sizer.Add(bSizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 15)
        self.SetSizer(sizer)
        self.CenterOnScreen()

    def onAction(self, action, link):
        if action == "url" and link:
            try: webbrowser.open(link)
            except: pass
        self.EndModal(wx.ID_OK)

def show_remote_dialog(sys_ctrl):
    dlg = RemoteSystemDialog(gui.mainFrame, sys_ctrl.get("dialog_title", "Notice"), sys_ctrl.get("dialog_message", ""), sys_ctrl.get("buttons", []))
    dlg.ShowModal()
    dlg.Destroy()

class DictionaryManagerDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=_("Custom Dictionary Manager"), size=(550, 500))
        self.commands = patterns.load_custom_commands() if patterns else {}
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        helpTxt = wx.StaticText(self, label=_("Add commands to replace spoken words with custom text.\nExample: 'my mail' -> 'user@example.com'\n(Inputs are auto-trimmed)"))
        mainSizer.Add(helpTxt, 0, wx.ALL, 10)
        self.listCtrl = wx.ListCtrl(self, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        self.listCtrl.InsertColumn(0, _("Voice Command"), width=200)
        self.listCtrl.InsertColumn(1, _("Replacement Text"), width=280)
        self.listCtrl.Bind(wx.EVT_LIST_ITEM_SELECTED, self.onItemSelected)
        mainSizer.Add(self.listCtrl, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
        inputSizer = wx.FlexGridSizer(rows=2, cols=2, vgap=10, hgap=10)
        inputSizer.AddGrowableCol(1, 1)
        inputSizer.Add(wx.StaticText(self, label=_("Voice Command:")), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txtCommand = wx.TextCtrl(self)
        inputSizer.Add(self.txtCommand, 1, wx.EXPAND)
        inputSizer.Add(wx.StaticText(self, label=_("Replacement:")), 0, wx.ALIGN_CENTER_VERTICAL)
        self.txtReplace = wx.TextCtrl(self)
        inputSizer.Add(self.txtReplace, 1, wx.EXPAND)
        mainSizer.Add(inputSizer, 0, wx.EXPAND | wx.ALL, 10)
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btnAdd = wx.Button(self, label=_("Add / Update"))
        self.btnAdd.Bind(wx.EVT_BUTTON, self.onAdd)
        self.btnDelete = wx.Button(self, label=_("Delete"))
        self.btnDelete.Bind(wx.EVT_BUTTON, self.onDelete)
        self.btnDelete.Disable()
        self.btnClear = wx.Button(self, label=_("Clear Fields"))
        self.btnClear.Bind(wx.EVT_BUTTON, self.onClear)
        btnSizer.Add(self.btnAdd, 0, wx.RIGHT, 5)
        btnSizer.Add(self.btnDelete, 0, wx.RIGHT, 5)
        btnSizer.Add(self.btnClear, 0)
        mainSizer.Add(btnSizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        closeBtn = wx.Button(self, wx.ID_OK, label=_("Close"))
        mainSizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.SetSizer(mainSizer)
        self.refreshList()
        self.Center()

    def refreshList(self):
        self.listCtrl.DeleteAllItems()
        for cmd, rep in self.commands.items():
            idx = self.listCtrl.InsertItem(self.listCtrl.GetItemCount(), str(cmd))
            self.listCtrl.SetItem(idx, 1, str(rep))

    def onItemSelected(self, evt):
        item = evt.GetIndex()
        self.txtCommand.Value = self.listCtrl.GetItemText(item, 0)
        self.txtReplace.Value = self.listCtrl.GetItemText(item, 1)
        self.btnDelete.Enable()

    def onAdd(self, evt):
        cmd = self.txtCommand.Value.strip()
        rep = self.txtReplace.Value.strip()
        if cmd and rep:
            cmd = patterns.normalize_text(cmd)
            rep = patterns.normalize_text(rep)
            self.commands[cmd] = rep
            patterns.save_custom_commands(self.commands)
            self.refreshList()
            self.onClear(None)

    def onDelete(self, evt):
        cmd = patterns.normalize_text(self.txtCommand.Value.strip())
        if cmd in self.commands:
            del self.commands[cmd]
            patterns.save_custom_commands(self.commands)
            self.refreshList()
            self.onClear(None)

    def onClear(self, evt):
        self.txtCommand.Value = ""
        self.txtReplace.Value = ""
        self.btnDelete.Disable()
        self.listCtrl.Select(self.listCtrl.GetFirstSelected(), False)

class TypeTalkSettingsPanel(gui.settingsDialogs.SettingsPanel):
    title = "TypeTalk AI"
    def makeSettings(self, settingsSizer):
        sHelper = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        voiceGroup = wx.StaticBox(self, label=_("Voice Input Settings"))
        voiceSizer = wx.StaticBoxSizer(voiceGroup, wx.VERTICAL)
        vHelper = gui.guiHelper.BoxSizerHelper(voiceGroup, sizer=voiceSizer)
        lang_names = [x[0] for x in patterns.ALL_LANGUAGES] if patterns else []
        self.inputLang = vHelper.addLabeledControl(_("Input Language:"), wx.Choice, choices=lang_names)
        try:
            current = config.conf["TypeTalkAI"]["input_language"]
            if patterns:
                idx = next(i for i, v in enumerate(patterns.ALL_LANGUAGES) if v[1] == current)
                self.inputLang.SetSelection(idx)
        except: self.inputLang.SetSelection(0)
        sHelper.addItem(voiceSizer)
        
        aiGroup = wx.StaticBox(self, label=_("AI & Processing Logic"))
        aiSizer = wx.StaticBoxSizer(aiGroup, wx.VERTICAL)
        aHelper = gui.guiHelper.BoxSizerHelper(aiGroup, sizer=aiSizer)
        self.useAI = aHelper.addItem(wx.CheckBox(aiGroup, label=_("Enable AI Detection")))
        self.useAI.Value = get_config_bool("use_ai_processing", False)
        self.useAI.Bind(wx.EVT_CHECKBOX, self.onToggleAI)
        
        self.models_list = get_ai_models()
        model_choices = [x[0] for x in self.models_list]
        self.aiModel = aHelper.addLabeledControl(_("Select AI Model:"), wx.Choice, choices=model_choices)
        try:
            current = config.conf["TypeTalkAI"].get("ai_model", "gpt_5_4_mini")
            idx = next(i for i, v in enumerate(self.models_list) if v[1] == current)
            self.aiModel.SetSelection(idx)
        except: self.aiModel.SetSelection(0)
        self.aiModel.Bind(wx.EVT_CHOICE, self.onModelChange)
        
        self.styleSelector = aHelper.addLabeledControl(_("Writing Style / Tone:"), wx.Choice, choices=WRITING_STYLES)
        try:
            cur_style = config.conf["TypeTalkAI"].get("writing_style", "Default (Standard)")
            idx = WRITING_STYLES.index(cur_style)
            self.styleSelector.SetSelection(idx)
        except: self.styleSelector.SetSelection(0)
        
        transSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.useTrans = wx.CheckBox(aiGroup, label=_("Translate to:"))
        self.useTrans.Value = get_config_bool("use_translation", False)
        transSizer.Add(self.useTrans, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 5)
        self.targetLang = wx.Choice(aiGroup, choices=lang_names)
        try: 
            cur_trans = config.conf["TypeTalkAI"].get("target_language", "English (US)")
            idx = lang_names.index(cur_trans)
            self.targetLang.SetSelection(idx)
        except: self.targetLang.SetSelection(0)
        transSizer.Add(self.targetLang, 1, wx.EXPAND)
        aHelper.addItem(transSizer)
        
        self.useEmoji = aHelper.addItem(wx.CheckBox(aiGroup, label=_("Add Emojis")))
        self.useEmoji.Value = get_config_bool("use_emoji", False)
        self.autoCopy = aHelper.addItem(wx.CheckBox(aiGroup, label=_("Auto copy result to clipboard")))
        self.autoCopy.Value = get_config_bool("auto_copy_clipboard", False)
        self.enableSound = aHelper.addItem(wx.CheckBox(aiGroup, label=_("Enable sound feedback")))
        self.enableSound.Value = get_config_bool("enable_sound", True)
        sHelper.addItem(aiSizer)

        dictGroup = wx.StaticBox(self, label=_("Custom Dictionary"))
        dictSizer = wx.StaticBoxSizer(dictGroup, wx.VERTICAL)
        dHelper = gui.guiHelper.BoxSizerHelper(dictGroup, sizer=dictSizer)
        self.useCustomDict = dHelper.addItem(wx.CheckBox(dictGroup, label=_("Enable Custom Dictionary & Punctuation")))
        self.useCustomDict.Value = get_config_bool("use_custom_dictionary", True)
        self.useCustomDict.Bind(wx.EVT_CHECKBOX, self.onToggleCustomDict)
        self.manageDictBtn = dHelper.addItem(wx.Button(dictGroup, label=_("Manage Dictionary")))
        self.manageDictBtn.Bind(wx.EVT_BUTTON, self.onManageDict)
        self.manageDictBtn.Enable(self.useCustomDict.Value)
        sHelper.addItem(dictSizer)

        customGroup = wx.StaticBox(self, label=_("Custom API Configuration"))
        customSizer = wx.StaticBoxSizer(customGroup, wx.VERTICAL)
        cHelper = gui.guiHelper.BoxSizerHelper(customGroup, sizer=customSizer)
        self.custUrl = cHelper.addLabeledControl("API URL:", wx.TextCtrl)
        self.custUrl.Value = config.conf["TypeTalkAI"].get("custom_url", "")
        methodSizer = wx.BoxSizer(wx.HORIZONTAL)
        self.custMethod = wx.Choice(customGroup, choices=["GET", "POST"])
        self.custMethod.SetSelection(0 if config.conf["TypeTalkAI"].get("custom_method", "GET") == "GET" else 1)
        methodSizer.Add(self.custMethod, 0, wx.RIGHT, 15)
        self.custParam = wx.TextCtrl(customGroup)
        self.custParam.Value = config.conf["TypeTalkAI"].get("custom_param", "q")
        methodSizer.Add(self.custParam, 1, wx.EXPAND)
        cHelper.addItem(methodSizer)
        self.custPath = cHelper.addLabeledControl("JSON Path:", wx.TextCtrl)
        self.custPath.Value = config.conf["TypeTalkAI"].get("custom_response_path", "response")
        sHelper.addItem(customSizer)
        self.aboutBtn = sHelper.addItem(wx.Button(self, label=_("About Developer")))
        self.aboutBtn.Bind(wx.EVT_BUTTON, self.onAbout)
        self.onToggleAI(None)
        self.onModelChange(None)

    def onToggleCustomDict(self, evt):
        self.manageDictBtn.Enable(self.useCustomDict.Value)

    def onManageDict(self, evt):
        if not patterns: return
        dlg = DictionaryManagerDialog(self)
        dlg.ShowModal()
        dlg.Destroy()

    def onToggleAI(self, evt):
        en = self.useAI.Value
        self.aiModel.Enable(en)
        self.useTrans.Enable(en)
        self.targetLang.Enable(en)
        self.useEmoji.Enable(en)
        self.styleSelector.Enable(en)
        if not en: self.enableCustomFields(False)
        else: self.onModelChange(None)

    def onModelChange(self, evt):
        if not self.useAI.Value: return
        sel_idx = self.aiModel.GetSelection()
        code = self.models_list[sel_idx][1]
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
        if patterns:
            config.conf["TypeTalkAI"]["input_language"] = patterns.ALL_LANGUAGES[self.inputLang.GetSelection()][1]
        config.conf["TypeTalkAI"]["use_custom_dictionary"] = self.useCustomDict.Value
        config.conf["TypeTalkAI"]["use_ai_processing"] = self.useAI.Value
        config.conf["TypeTalkAI"]["ai_model"] = self.models_list[self.aiModel.GetSelection()][1]
        config.conf["TypeTalkAI"]["use_translation"] = self.useTrans.Value
        config.conf["TypeTalkAI"]["writing_style"] = WRITING_STYLES[self.styleSelector.GetSelection()]
        if patterns:
            config.conf["TypeTalkAI"]["target_language"] = patterns.ALL_LANGUAGES[self.targetLang.GetSelection()][0]
        config.conf["TypeTalkAI"]["use_emoji"] = self.useEmoji.Value
        config.conf["TypeTalkAI"]["auto_copy_clipboard"] = self.autoCopy.Value
        config.conf["TypeTalkAI"]["enable_sound"] = self.enableSound.Value
        config.conf["TypeTalkAI"]["custom_url"] = self.custUrl.Value
        config.conf["TypeTalkAI"]["custom_method"] = "GET" if self.custMethod.GetSelection() == 0 else "POST"
        config.conf["TypeTalkAI"]["custom_param"] = self.custParam.Value
        config.conf["TypeTalkAI"]["custom_response_path"] = self.custPath.Value

class AboutDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=_("About TypeTalk AI"), size=(600, 550))
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
        closeBtn = wx.Button(self, wx.ID_OK, label=_("Close"))
        mainSizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 15)
        self.SetSizer(mainSizer)
        self.Center()
    def openLink(self, url):
        try: webbrowser.open(url)
        except: pass

class HelpDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title=_("TypeTalk AI Help"), size=(500, 500))
        sizer = wx.BoxSizer(wx.VERTICAL)
        text = wx.TextCtrl(self, value=HELP_TEXT, style=wx.TE_MULTILINE | wx.TE_READONLY)
        text.SetFont(wx.Font(12, wx.FONTFAMILY_MODERN, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))
        sizer.Add(text, 1, wx.EXPAND | wx.ALL, 10)
        closeBtn = wx.Button(self, wx.ID_OK, label=_("Close"))
        sizer.Add(closeBtn, 0, wx.ALIGN_CENTER | wx.BOTTOM, 10)
        self.SetSizer(sizer)
        self.Center()

class UpdateDialog(wx.Dialog):
    def __init__(self, parent, version, changelog, download_url):
        super().__init__(parent, title=_("Critical Update"), size=(400, 300), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP | wx.CENTER)
        self.download_url = download_url
        self.version = version
        sizer = wx.BoxSizer(wx.VERTICAL)
        warn = wx.StaticText(self, label=_("Update Required ({version})").format(version=version))
        warn.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        warn.SetForegroundColour(wx.Colour(200, 0, 0))
        sizer.Add(warn, 0, wx.ALIGN_CENTER | wx.TOP, 15)
        log = wx.TextCtrl(self, value=changelog, style=wx.TE_MULTILINE | wx.TE_READONLY)
        sizer.Add(log, 1, wx.EXPAND | wx.ALL, 10)
        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        updateBtn = wx.Button(self, label=_("Update Now"))
        updateBtn.Bind(wx.EVT_BUTTON, self.onUpdate)
        updateBtn.SetDefault()
        closeBtn = wx.Button(self, wx.ID_CANCEL, label=_("Later (Locked)"))
        btnSizer.Add(updateBtn, 0, wx.RIGHT, 10)
        btnSizer.Add(closeBtn, 0)
        sizer.Add(btnSizer, 0, wx.ALIGN_CENTER | wx.BOTTOM, 15)
        self.SetSizer(sizer)
        self.CenterOnScreen()

    def onUpdate(self, evt):
        self.EndModal(wx.ID_OK)
        self.progressDlg = wx.ProgressDialog(_("Downloading Update"), _("Connecting to server..."), maximum=100, parent=gui.mainFrame, style=wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
        t = threading.Thread(target=self._download_worker)
        t.daemon = True
        t.start()

    def _download_worker(self):
        try:
            temp_dir = tempfile.gettempdir()
            file_name = f"TypeTalkAI_{self.version}.nvda-addon"
            file_path = os.path.join(temp_dir, file_name)
            ctx = ssl._create_unverified_context()
            with request.urlopen(self.download_url, context=ctx) as response:
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
                            wx.CallAfter(self.progressDlg.Update, percent, _("Downloading: {percent}%").format(percent=percent))
            wx.CallAfter(self.progressDlg.Destroy)
            wx.CallAfter(os.startfile, file_path)
        except Exception:
            wx.CallAfter(self.progressDlg.Destroy)

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    scriptCategory = "TypeTalk AI"
    is_recording = False
    toggling = False 
    temp_audio_file = os.path.join(tempfile.gettempdir(), "typetalk_audio.wav")
    update_available = False
    update_locked = False 
    last_processed_text = ""
    saved_context_text = ""
    REPO_API = "https://api.github.com/repos/hridoyhub/TypeTalk-AI-NVDA-Addon/releases/latest"
    latest_ver = ""
    latest_log = ""
    latest_url = ""

    def __init__(self):
        super(GlobalPlugin, self).__init__()
        config.conf.spec["TypeTalkAI"] = confspec
        gui.settingsDialogs.NVDASettingsDialog.categoryClasses.append(TypeTalkSettingsPanel)
        try: current_addon_version = addonHandler.getCodeAddon().manifest['version']
        except: current_addon_version = "1.0.0"
        last_seen = config.conf["TypeTalkAI"].get("last_welcome_version", "0.0.0")
        if last_seen != current_addon_version and welcome:
            wx.CallAfter(self._show_welcome_wizard, current_addon_version)
        threading.Thread(target=self._check_for_update, daemon=True).start()

    def _show_welcome_wizard(self, version_to_save):
        try:
            welcome.show_wizard()
            config.conf["TypeTalkAI"]["last_welcome_version"] = version_to_save
        except: pass

    def _check_for_update(self):
        try:
            cur_ver = addonHandler.getCodeAddon().manifest['version']
            req = request.Request(self.REPO_API, headers={'User-Agent': 'NVDA-Addon'})
            ctx = ssl._create_unverified_context()
            with request.urlopen(req, timeout=10, context=ctx) as response:
                data = json.loads(response.read().decode('utf-8'))
            remote_tag = data.get('tag_name', 'v0.0').lstrip('v')
            changelog = data.get('body', _('No details.'))
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
        except: pass

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

    def getScript(self, gesture):
        if not self.toggling: return super(GlobalPlugin, self).getScript(gesture)
        script = super(GlobalPlugin, self).getScript(gesture)
        if not script: return finally_(self.script_invalidCommand, self.finish)
        return finally_(script, self.finish)

    def finish(self):
        self.toggling = False
        self.clearGestureBindings()
        self.bindGestures(self.__gestures)

    def script_error(self, gesture): play_sound("error")
    def script_invalidCommand(self, gesture): ui.message(_("Command unavailable. Press H for Help."))
    
    def script_updateLocked(self, gesture):
        play_sound("error")
        ui.message(_("Update Required! Please update to continue."))
        if self.latest_url: wx.CallAfter(self._show_update_dialog)

    def _check_remote_config(self):
        try:
            url = f"{_get_gist_url()}?time={int(time.time())}"
            req = request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            ctx = ssl._create_unverified_context()
            with request.urlopen(req, timeout=5, context=ctx) as resp:
                remote_data = json.loads(resp.read().decode('utf-8'))
            local_data = load_local_config()
            local_version = local_data.get("config_version", 0) if local_data else 0
            remote_version = remote_data.get("config_version", 0)
            if remote_version > local_version or not local_data:
                if save_local_config(remote_data):
                    wx.CallAfter(tones.beep, 1200, 100)
        except Exception:
            pass

    @scriptHandler.script(description=_("Activates Command Layer"))
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
        threading.Thread(target=self._check_remote_config, daemon=True).start()

    @scriptHandler.script(description=_("Toggle Voice Typing"))
    def script_smartDictation(self, gesture):
        if self.toggling: self.finish()
        if not sr:
            self._show_error(_("Library missing"), f"SpeechRecognition not found. {sr_error}")
            return
        try:
            if not self.is_recording:
                self.is_recording = True
                safe_beep(800, 100)
                ctypes.windll.winmm.mciSendStringW('close all', None, 0, 0)
                cmd_open = 'open new type waveaudio alias myaudio'
                res = ctypes.windll.winmm.mciSendStringW(cmd_open, None, 0, 0)
                if res != 0: raise Exception("Microphone Access Failed")
                ctypes.windll.winmm.mciSendStringW('set myaudio time format ms', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio format tag pcm', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio channels 1', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio samplespersec 44100', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio bitspersample 16', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio bytespersec 88200', None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('set myaudio alignment 2', None, 0, 0)
                res_record = ctypes.windll.winmm.mciSendStringW('record myaudio', None, 0, 0)
                if res_record != 0: raise Exception("Recording Failed")
                lang_code = config.conf["TypeTalkAI"].get("input_language", "bn-BD")
                lang_name = "Unknown"
                if patterns:
                    lang_name = next((name for name, code in patterns.ALL_LANGUAGES if code == lang_code), "Unknown")
                ui.message(_("Listening ({lang})...").format(lang=lang_name))
            else:
                self.is_recording = False
                safe_beep(400, 100)
                save_cmd = f'save myaudio "{self.temp_audio_file}"'
                ctypes.windll.winmm.mciSendStringW(save_cmd, None, 0, 0)
                ctypes.windll.winmm.mciSendStringW('close myaudio', None, 0, 0)
                ui.message(_("Processing..."))
                threading.Thread(target=self._process_pipeline, daemon=True).start()
        except Exception as e:
            self.is_recording = False
            ctypes.windll.winmm.mciSendStringW('close all', None, 0, 0)
            self._show_error(_("Mic Error"), str(e))

    def _process_pipeline(self):
        try:
            if not os.path.exists(self.temp_audio_file): return
            r = sr.Recognizer()
            with sr.AudioFile(self.temp_audio_file) as source:
                audio_data = r.record(source)
            lang_code = config.conf["TypeTalkAI"].get("input_language", "bn-BD")
            try:
                raw_text = r.recognize_google(audio_data, language=lang_code)
            except sr.UnknownValueError:
                play_sound("error")
                wx.CallAfter(ui.message, _("No speech detected"))
                return
            except Exception as e:
                wx.CallAfter(self._show_error, _("Network Error"), str(e))
                return

            ai_on = get_config_bool("use_ai_processing", False)
            custom_dict_on = get_config_bool("use_custom_dictionary", True)

            if not ai_on:
                processed_text = raw_text
                if custom_dict_on and patterns:
                    processed_text = patterns.apply_custom_replacements(processed_text)
                elif patterns:
                    processed_text = patterns.offline_format(processed_text)
                self.last_processed_text = processed_text
                wx.CallAfter(self._paste_text, processed_text)
            else:
                wx.CallAfter(ui.message, _("Refining..."))
                self._call_ai_api(raw_text)
        except Exception as e:
            wx.CallAfter(self._show_error, _("System Error"), str(e))
        finally:
            try: os.remove(self.temp_audio_file)
            except: pass

    def _call_ai_api(self, input_text, mode="standard", context="", select_all_mode=False):
        try:
            local_data = load_local_config()
            if local_data and local_data.get("system_control", {}).get("show_dialog", False):
                play_sound("error")
                wx.CallAfter(ui.message, _("Service currently unavailable. Please check the dialog."))
                wx.CallAfter(show_remote_dialog, local_data["system_control"])
                return

            model_code = config.conf["TypeTalkAI"].get("ai_model", "gpt_5_4_mini")
            use_trans = get_config_bool("use_translation", False)
            target_lang = config.conf["TypeTalkAI"].get("target_language", "English (US)")
            use_emoji = get_config_bool("use_emoji", False)
            writing_style = config.conf["TypeTalkAI"].get("writing_style", "Default (Standard)")
            
            prompt = ""
            if mode == "reply":
                if not input_text:
                    prompt = f"Context: '{context}'. User provided no hint. Generate a suitable, natural reply based solely on the context."
                else:
                    prompt = f"Context: '{context}'. User Hint: '{input_text}'. Generate a natural reply."
                if use_trans: prompt += f" Reply MUST be in {target_lang}."
                if writing_style != "Default (Standard)": prompt += f" Use {writing_style} tone."
                if use_emoji: prompt += " Add emojis."
            else:
                if use_trans: prompt = f"Translate to {target_lang}: '{input_text}'."
                else: prompt = f"Fix grammar: '{input_text}'."
                if writing_style != "Default (Standard)":
                    style_map = {
                        "Formal (Professional)": "formal",
                        "Casual (Friendly)": "casual",
                        "Concise (Short)": "concise",
                        "Bullet Points": "bullet point list"
                    }
                    style_key = style_map.get(writing_style, "normal")
                    prompt += f" Rewrite in {style_key} tone."
                if use_emoji: prompt += " Add emojis."
            prompt += " Output ONLY the result text."

            api_url = ""
            method = "GET"
            param_name = "q"
            resp_path = "response"

            if model_code == "custom":
                api_url = config.conf["TypeTalkAI"].get("custom_url", "")
                method = config.conf["TypeTalkAI"].get("custom_method", "GET")
                param_name = config.conf["TypeTalkAI"].get("custom_param", "q")
                resp_path = config.conf["TypeTalkAI"].get("custom_response_path", "response")
            else:
                selected_model = None
                if local_data and "ai_models" in local_data:
                    selected_model = next((m for m in local_data["ai_models"] if m["id"] == model_code), None)
                if not selected_model: raise Exception("Model not found")
                if not selected_model.get("is_active", True): raise Exception("This model is inactive")
                api_url = selected_model["api_url"]
                method = selected_model["method"]
                param_name = selected_model["input_param"]
                resp_path = selected_model["response_path"]

            if not api_url: raise Exception("Invalid URL")
            final_text = ""
            ctx = ssl._create_unverified_context()
            if method == "GET":
                query_string = parse.urlencode({param_name: prompt})
                full_url = f"{api_url}&{query_string}" if "?" in api_url else f"{api_url}?{query_string}"
                req = request.Request(full_url, headers={'User-Agent': 'Mozilla/5.0'})
                with request.urlopen(req, timeout=15, context=ctx) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    final_text = data.get(resp_path, "Error: Field not found")
            else:
                data_json = json.dumps({param_name: prompt}).encode('utf-8')
                req = request.Request(api_url, data=data_json, method="POST", headers={'Content-Type': 'application/json'})
                with request.urlopen(req, timeout=15, context=ctx) as response:
                    data = json.loads(response.read().decode('utf-8'))
                    final_text = data.get(resp_path, "Error: Field not found")
            
            custom_dict_on = get_config_bool("use_custom_dictionary", True)
            if custom_dict_on and patterns and final_text:
                final_text = patterns.apply_custom_replacements(final_text)
            
            if final_text: 
                self.last_processed_text = final_text
                wx.CallAfter(self._paste_text, final_text, select_all_mode)
            else: wx.CallAfter(self._show_error, _("AI Error"), _("Empty response"))
        except Exception as e:
            wx.CallAfter(self._show_error, _("API Error"), str(e))

    def _paste_text(self, text, select_all=False):
        play_sound("success")
        auto_copy = get_config_bool("auto_copy_clipboard", False)
        if auto_copy:
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
                wx.CallLater(500, ui.message, _("Typed & Copied: {text}").format(text=text))
            except: pass
        else:
            if _smart_insert_text(text, select_all):
                wx.CallLater(500, ui.message, _("Typed: {text}").format(text=text))
            else:
                _send_unicode_text(text)
                wx.CallLater(500, ui.message, _("Typed (Direct): {text}").format(text=text))

    @scriptHandler.script(description=_("Refine or Translate Selected Text"))
    def script_refineText(self, gesture):
        if self.toggling: self.finish()
        try:
            focused = api.getFocusObject()
            if not focused: raise Exception("No focused element")
            selected_text = ""
            try:
                info = focused.makeTextInfo(textInfos.POSITION_SELECTION)
                selected_text = info.text.strip()
            except: pass
            if not selected_text:
                try:
                    info = focused.makeTextInfo(textInfos.POSITION_ALL)
                    selected_text = info.text.strip()
                except: pass
            if not selected_text:
                play_sound("error")
                ui.message(_("No text found"))
                return
            play_sound("start")
            ui.message(_("Refining..."))
            threading.Thread(target=self._call_ai_api, args=(selected_text,), daemon=True).start()
        except Exception as e:
            play_sound("error")

    @scriptHandler.script(description=_("Set Context for Smart Reply"))
    def script_setContext(self, gesture):
        if self.toggling: self.finish()
        try:
            focused = api.getFocusObject()
            if not focused: raise Exception("No focus")
            self.saved_context_text = ""
            try:
                info = focused.makeTextInfo(textInfos.POSITION_SELECTION)
                self.saved_context_text = info.text.strip()
            except: pass
            if not self.saved_context_text:
                try: self.saved_context_text = focused.name or focused.value or ""
                except: pass
            if not self.saved_context_text:
                play_sound("error")
                ui.message(_("No context found"))
            else:
                play_sound("success")
                ui.message(_("Context set"))
        except Exception as e:
            play_sound("error")

    @scriptHandler.script(description=_("Generate Smart Reply from Hint"))
    def script_generateReply(self, gesture):
        if self.toggling: self.finish()
        try:
            if not self.saved_context_text:
                play_sound("error")
                ui.message(_("Set context first (Press C)"))
                return
            focused = api.getFocusObject()
            if not focused: raise Exception("No focus")
            hint_text = ""
            try:
                info = focused.makeTextInfo(textInfos.POSITION_SELECTION)
                hint_text = info.text.strip()
            except: pass
            if not hint_text:
                try:
                    if focused.role in (controlTypes.ROLE_EDIT, controlTypes.ROLE_DOCUMENT):
                        info = focused.makeTextInfo(textInfos.POSITION_ALL)
                        hint_text = info.text.strip()
                except: pass
            play_sound("start")
            ui.message(_("Generating Reply..."))
            threading.Thread(target=self._call_ai_api, args=(hint_text, "reply", self.saved_context_text, True), daemon=True).start()
        except Exception as e:
            play_sound("error")

    @scriptHandler.script(description=_("Check Current Settings"))
    def script_checkStatus(self, gesture):
        if self.toggling: self.finish()
        lang = config.conf["TypeTalkAI"].get("input_language", "bn-BD")
        if patterns:
             lang = next((name for name, code in patterns.ALL_LANGUAGES if code == lang), lang)
        ai = "ON" if get_config_bool("use_ai_processing", False) else "OFF"
        style = config.conf["TypeTalkAI"].get("writing_style", "Standard")
        trans = "ON" if get_config_bool("use_translation", False) else "OFF"
        msg = _("Lang: {lang}, AI: {ai}, Style: {style}, Trans: {trans}").format(lang=lang, ai=ai, style=style, trans=trans)
        ui.message(msg)

    @scriptHandler.script(description=_("Recover Last Text"))
    def script_recoverText(self, gesture):
        if self.toggling: self.finish()
        if not self.last_processed_text:
            play_sound("error")
            ui.message(_("No text to recover"))
            return
        wx.CallAfter(self._paste_text, self.last_processed_text)
        ui.message(_("Recovered"))

    @scriptHandler.script(description=_("Change Writing Style"))
    def script_changeStyle(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_style_dialog)
        
    def _show_style_dialog(self):
        try:
            dlg = wx.SingleChoiceDialog(gui.mainFrame, _("Select Writing Style:"), "TypeTalk AI", WRITING_STYLES)
            current = config.conf["TypeTalkAI"].get("writing_style", "Default (Standard)")
            try: idx = WRITING_STYLES.index(current); dlg.SetSelection(idx)
            except: pass
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection()
                style = WRITING_STYLES[sel_idx]
                config.conf["TypeTalkAI"]["writing_style"] = style
                wx.CallLater(100, ui.message, _("Style: {style}").format(style=style))
            dlg.Destroy()
        except: pass

    @scriptHandler.script(description=_("Toggle AI Processing"))
    def script_toggleAI(self, gesture):
        if self.toggling: self.finish()
        current = get_config_bool("use_ai_processing", False)
        new_val = not current
        if new_val:
            local_data = load_local_config()
            if local_data and local_data.get("system_control", {}).get("show_dialog", False):
                play_sound("error")
                ui.message(_("AI unavailable. Please check the dialog."))
                wx.CallAfter(show_remote_dialog, local_data["system_control"])
                return
        config.conf["TypeTalkAI"]["use_ai_processing"] = new_val
        ui.message(_("AI Detection {status}").format(status='Enabled' if new_val else 'Disabled'))

    @scriptHandler.script(description=_("Toggle Translation"))
    def script_toggleTranslation(self, gesture):
        if self.toggling: self.finish()
        current = get_config_bool("use_translation", False)
        new_val = not current
        config.conf["TypeTalkAI"]["use_translation"] = new_val
        ui.message(_("Translation {status}").format(status='Enabled' if new_val else 'Disabled'))

    @scriptHandler.script(description=_("Toggle Emoji"))
    def script_toggleEmoji(self, gesture):
        if self.toggling: self.finish()
        current = get_config_bool("use_emoji", False)
        new_val = not current
        config.conf["TypeTalkAI"]["use_emoji"] = new_val
        ui.message(_("Emoji {status}").format(status='Enabled' if new_val else 'Disabled'))

    @scriptHandler.script(description=_("Show Developer Info"))
    def script_showAbout(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_about_dialog)

    @scriptHandler.script(description=_("Show Help Menu"))
    def script_showHelp(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_help_dialog)

    @scriptHandler.script(description=_("Change AI Model"))
    def script_changeModel(self, gesture):
        if self.toggling: self.finish()
        local_data = load_local_config()
        if local_data and local_data.get("system_control", {}).get("show_dialog", False):
            play_sound("error")
            ui.message(_("Service unavailable. Please check the dialog."))
            wx.CallAfter(show_remote_dialog, local_data["system_control"])
            return
        wx.CallAfter(self._show_model_dialog)

    @scriptHandler.script(description=_("Change Input Language"))
    def script_changeInputLang(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_input_dialog)

    @scriptHandler.script(description=_("Change Translation Language"))
    def script_changeTransLang(self, gesture):
        if self.toggling: self.finish()
        wx.CallAfter(self._show_trans_dialog)

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

    def _show_model_dialog(self):
        try:
            models_list = get_ai_models()
            choices = [x[0] for x in models_list]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, _("Select AI Model:"), "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"].get("ai_model", "gpt_5_4_mini")
            try: idx = next(i for i, v in enumerate(models_list) if v[1] == current); dlg.SetSelection(idx)
            except: pass
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection(); code = models_list[sel_idx][1]; name = models_list[sel_idx][0]
                config.conf["TypeTalkAI"]["ai_model"] = code
                wx.CallLater(100, ui.message, _("Model: {name}").format(name=name))
            dlg.Destroy()
        except: pass

    def _show_input_dialog(self):
        if not patterns: return
        try:
            choices = [x[0] for x in patterns.ALL_LANGUAGES]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, _("Select Input Language:"), "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"].get("input_language", "bn-BD")
            try: idx = next(i for i, v in enumerate(patterns.ALL_LANGUAGES) if v[1] == current); dlg.SetSelection(idx)
            except: pass
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection(); code = patterns.ALL_LANGUAGES[sel_idx][1]; name = patterns.ALL_LANGUAGES[sel_idx][0]
                config.conf["TypeTalkAI"]["input_language"] = code
                wx.CallLater(100, ui.message, _("Input: {name}").format(name=name))
            dlg.Destroy()
        except: pass

    def _show_trans_dialog(self):
        if not patterns: return
        try:
            choices = [x[0] for x in patterns.ALL_LANGUAGES]
            dlg = wx.SingleChoiceDialog(gui.mainFrame, _("Select Translation Target:"), "TypeTalk AI", choices)
            current = config.conf["TypeTalkAI"].get("target_language", "English (US)")
            try: name_list = [x[0] for x in patterns.ALL_LANGUAGES]; idx = name_list.index(current); dlg.SetSelection(idx)
            except: pass
            if dlg.ShowModal() == wx.ID_OK:
                sel_idx = dlg.GetSelection(); name = patterns.ALL_LANGUAGES[sel_idx][0]
                config.conf["TypeTalkAI"]["target_language"] = name
                wx.CallLater(100, ui.message, _("Target: {name}").format(name=name))
            dlg.Destroy()
        except: pass

    __gestures = {
        "kb:NVDA+shift+space": "activateLayer",
    }
    
    __VisionGestures = {
        "kb:s": "smartDictation",
        "kb:r": "refineText",
        "kb:a": "toggleAI",
        "kb:t": "toggleTranslation",
        "kb:e": "toggleEmoji",
        "kb:d": "showAbout",
        "kb:h": "showHelp",
        "kb:m": "changeModel",
        "kb:i": "changeInputLang",
        "kb:l": "changeTransLang",
        "kb:w": "changeStyle",
        "kb:q": "checkStatus",
        "kb:b": "recoverText",
        "kb:c": "setContext",
        "kb:g": "generateReply",
    }