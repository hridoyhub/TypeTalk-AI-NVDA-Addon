# -*- coding: utf-8 -*-
import re
import json
import os
import logging
import unicodedata

# --- 1. Language Dictionary (All Languages Supported) ---
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

# --- 2. Custom Dictionary Management ---
DICT_FILE = os.path.join(os.path.dirname(__file__), "custom_commands.json")

def normalize_text(text):
    """Normalizes text to NFC form to handle Unicode (Bengali/Nepali/Hindi) correctly."""
    if not text: return ""
    return unicodedata.normalize('NFC', text)

def load_custom_commands():
    """Loads custom commands securely with UTF-8 and Normalization."""
    if not os.path.exists(DICT_FILE):
        return {}
    try:
        with open(DICT_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                # Normalize keys and values immediately upon load
                normalized_data = {normalize_text(k): normalize_text(v) for k, v in data.items()}
                return normalized_data
            return {}
    except Exception as e:
        logging.error(f"TypeTalkAI: Failed to load dictionary: {e}")
        return {}

def save_custom_commands(commands):
    """Saves commands ensuring ASCII characters are not escaped (readable text)."""
    try:
        # Normalize before saving
        normalized_commands = {normalize_text(k): normalize_text(v) for k, v in commands.items()}
        with open(DICT_FILE, "w", encoding="utf-8") as f:
            json.dump(normalized_commands, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logging.error(f"TypeTalkAI: Failed to save dictionary: {e}")

def apply_custom_replacements(text):
    """Applies replacements using JSON data with smart word boundaries."""
    if not text: return ""
    
    # 1. Normalize Input Text
    text = normalize_text(text)
    
    commands = load_custom_commands()
    
    # Universal Fallback: If no commands exist, return text as is
    if not commands: return text

    # Sort by length to match longer phrases first
    sorted_keys = sorted(commands.keys(), key=len, reverse=True)

    for cmd in sorted_keys:
        clean_cmd = cmd.strip()
        replacement = commands[cmd]
        
        # Robust Regex for Word Boundaries (Fixes Bengali/Nepali/Hindi issues)
        pattern = r"(?<!\w)" + re.escape(clean_cmd) + r"(?!\w)"
        
        try:
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE | re.UNICODE)
        except:
            text = text.replace(clean_cmd, replacement)
            
    # --- Final Cleanup (Smart Spacing for Punctuation) ---
    text = re.sub(r'\s+([,??!:;])', r'\1', text) 
    text = re.sub(r'([,??!:;])(?=[^\s\n])', r'\1 ', text)
    
    # Capitalize first letter (Universal safe check)
    text = text.strip()
    if text: 
        if text[0].islower():
            text = text[0].upper() + text[1:]
    
    return text

def offline_format(text):
    """
    Basic formatting for offline text when Custom Dict is OFF.
    Designed to be safe for all languages (Universal Pass-through).
    """
    text = text.strip()
    if not text: return ""
    
    # Capitalize first letter
    if text[0].islower():
        text = text[0].upper() + text[1:]
        
    # Only add punctuation if it's likely a sentence and missing one.
    if text[-1] not in ".!?|?": 
        text += "."
        
    return text