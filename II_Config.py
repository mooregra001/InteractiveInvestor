import json
import os
from II_Constants import DEFAULT_BASE_PATH, DEFAULT_EXCEL_PATH

def load_config(config_path="config.json"):
    """Load configuration from a JSON file.
    
    Args:
        config_path (str): Path to the configuration file. Defaults to 'config.json'.
    
    Returns:
        dict: Configuration dictionary with 'base_path' and 'excel_path' keys.
    
    Raises:
        json.JSONDecodeError: If the JSON file is invalid.
    """
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
            # Ensure base_path and excel_path are set, using defaults if missing
            config.setdefault("base_path", DEFAULT_BASE_PATH)
            config.setdefault("excel_path", DEFAULT_EXCEL_PATH)
            return config
    except FileNotFoundError:
        print(f"Warning: Configuration file not found: {config_path}. Using default paths.")
        return {"base_path": DEFAULT_BASE_PATH, "excel_path": DEFAULT_EXCEL_PATH}
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON in configuration file: {config_path}")
        raise