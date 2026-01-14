# -*- coding: utf-8 -*-
import config

def onInstall():
    pass

def onUninstall():
    try:
        # Reset welcome screen flag on uninstall
        if "TypeTalkAI" in config.conf:
            config.conf["TypeTalkAI"]["welcome_shown"] = False
    except:
        pass