# -*- coding: utf-8 -*-
import wx
import webbrowser
import gui

# --- Content Definitions ---

TITLE_1 = "Welcome to TypeTalk AI"
TEXT_1 = (
    "Thank you for installing TypeTalk AI.\n"
    "You have taken an important step toward a smarter, faster, and more accessible typing experience.\n\n"
    "TypeTalk AI is more than a voice typing tool. It is an intelligent assistant built specifically for NVDA users. "
    "It works by capturing your voice through Google Speech Recognition technology and then enhancing it using advanced artificial intelligence.\n\n"
    "This process transforms your spoken words into clear, well-structured, and grammatically correct text."
)

TITLE_2 = "The Vision Behind the Code"
TEXT_2 = (
    "We believe technology should remove barriers, not create them.\n"
    "TypeTalk AI was created with a strong commitment to empowering the visually impaired community through modern AI-driven solutions.\n\n"
    "Our mission is simple: to make digital communication smooth, natural, and confidence-building for everyone. "
    "We aim to reduce the gap between thought and text, ensuring your ideas are captured exactly as you intend — without the frustration of traditional typing.\n\n"
    "With TypeTalk AI, we are redefining accessibility."
)

TITLE_3 = "Join Our Journey"
TEXT_3 = (
    "TypeTalk AI is a passion-driven project developed by Md Hridoy Sheikh from Dhaka, Bangladesh.\n\n"
    "This is only the beginning. The future of TypeTalk AI is shaped by its users. "
    "We are continuously improving the tool, and your feedback truly matters. We invite you to join our growing community to share suggestions, report issues, and help guide future development.\n\n"
    "Together, we can build a more inclusive and accessible digital world.\n\n"
    "Click 'Finish' to begin your journey with TypeTalk AI!"
)

LINKS = {
    "Hridoy Modding Hub": "https://t.me/Hridoy_Modding_Hub",
    "School of Mind Light": "https://t.me/schoolofmindlight2018",
    "Omnisent Community": "https://t.me/omnisent25",
    "Helpful App Store": "https://t.me/Helpfulappstore"
}

class WelcomeStep(wx.Dialog):
    def __init__(self, parent, title, text, step, total_steps):
        super().__init__(parent, title=title, size=(450, 400), style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP)
        
        self.step = step
        self.result_action = None

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        heading = wx.StaticText(self, label=title)
        heading.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        heading.SetForegroundColour(wx.Colour(0, 51, 102))
        mainSizer.Add(heading, 0, wx.ALIGN_CENTER | wx.TOP, 20)

        body = wx.TextCtrl(self, value=text, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_NO_VSCROLL | wx.BORDER_NONE)
        body.SetBackgroundColour(self.GetBackgroundColour())
        mainSizer.Add(body, 1, wx.EXPAND | wx.ALL, 20)

        if step == 3:
            linkSizer = wx.GridSizer(rows=2, cols=2, vgap=10, hgap=10)
            for name, url in LINKS.items():
                btn = wx.Button(self, label=name)
                btn.Bind(wx.EVT_BUTTON, lambda evt, u=url: webbrowser.open(u))
                linkSizer.Add(btn, 0, wx.EXPAND)
            mainSizer.Add(linkSizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 20)

        btnSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        stepLbl = wx.StaticText(self, label=f"Step {step} of {total_steps}")
        btnSizer.Add(stepLbl, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 20)

        if step < total_steps:
            nextBtn = wx.Button(self, label="Next >")
            nextBtn.Bind(wx.EVT_BUTTON, self.onNext)
            nextBtn.SetDefault()
            btnSizer.Add(nextBtn, 0)
        else:
            finishBtn = wx.Button(self, label="Finish")
            finishBtn.Bind(wx.EVT_BUTTON, self.onFinish)
            finishBtn.SetDefault()
            btnSizer.Add(finishBtn, 0)

        mainSizer.Add(btnSizer, 0, wx.ALIGN_RIGHT | wx.ALL, 15)

        self.SetSizer(mainSizer)
        self.CenterOnScreen()

    def onNext(self, evt):
        self.result_action = "next"
        self.EndModal(wx.ID_OK)

    def onFinish(self, evt):
        self.result_action = "finish"
        self.EndModal(wx.ID_OK)

def show_wizard():
    dlg1 = WelcomeStep(gui.mainFrame, TITLE_1, TEXT_1, 1, 3)
    dlg1.ShowModal()
    action = dlg1.result_action
    dlg1.Destroy()
    
    if action != "next": return

    dlg2 = WelcomeStep(gui.mainFrame, TITLE_2, TEXT_2, 2, 3)
    dlg2.ShowModal()
    action = dlg2.result_action
    dlg2.Destroy()

    if action != "next": return

    dlg3 = WelcomeStep(gui.mainFrame, TITLE_3, TEXT_3, 3, 3)
    dlg3.ShowModal()
    dlg3.Destroy()