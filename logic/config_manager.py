import os
import json
from typing import Dict, Any

CONFIG_FILE = "file_organizer_config.json"


class ConfigManager:
    """Handles loading and saving of the application configuration."""

    @staticmethod
    def load_config() -> Dict[str, Any]:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                print(f"Error loading config file or file is corrupted: {CONFIG_FILE}")
                return {}
        return {}

    @staticmethod
    def save_config(config: Dict[str, Any]) -> None:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")
