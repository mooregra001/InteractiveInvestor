# II_Config.py
import json
import os
from II_Constants import DEFAULT_BASE_PATH, DEFAULT_EXCEL_PATH

def load_config(config_path="config.json"):
    """Load configuration from JSON file with defaults."""
    default_cfg = {
        "base_path": DEFAULT_BASE_PATH,
        "excel_path": DEFAULT_EXCEL_PATH
    }

    if not os.path.exists(config_path):
        print(f"Warning: Config file not found: {config_path}. Using defaults and creating file.")
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_cfg, f, indent=4)
        return default_cfg

    try:
        with open(config_path, 'r', encoding="utf-8") as f:
            config = json.load(f)
        # Merge defaults
        for key, value in default_cfg.items():
            config.setdefault(key, value)
        return config
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON in {config_path}. Using defaults.")
        return default_cfg
    except Exception as e:
        print(f"Error loading config: {e}")
        return default_cfg


def update_config(new_excel_path, config_path="config.json"):
    """Update config.json with new excel_path."""
    config = load_config(config_path)
    config["excel_path"] = new_excel_path
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4)
        print(f"Config updated: excel_path = {new_excel_path}")
    except Exception as e:
        raise RuntimeError(f"Failed to write config: {e}")