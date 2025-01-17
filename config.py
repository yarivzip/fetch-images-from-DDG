import json
import os
import logging

DEFAULT_CONFIG = {
    "description_column": "תאור",
    "filename_column": "שם קובץ",
    "max_size": "800",
    "concurrent_downloads": "3",
    "skip_existing": True,
    "download_directory": "downloaded_images"  # Default download directory
}

CONFIG_FILE = "user_preferences.json"

def load_config():
    """Load user preferences from JSON file"""
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # Merge with defaults to ensure all keys exist
                return {**DEFAULT_CONFIG, **config}
        else:
            logging.info("No config file found, using defaults")
            return DEFAULT_CONFIG.copy()
    except Exception as e:
        logging.error(f"Error loading config: {e}")
        return DEFAULT_CONFIG.copy()

def save_config(config):
    """Save user preferences to JSON file"""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        logging.info("Configuration saved successfully")
    except Exception as e:
        logging.error(f"Error saving config: {e}")
