import configparser
from pathlib import Path
import logging
import sys

CONFIG = None

def _get_base_path():
    """Get the base path of the application."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)  # PyInstaller temp dir
    return Path(__file__).parent

def load_config(config_name="config.ini"):
    """Read config.ini from the package directory."""
    global CONFIG
    base_path = _get_base_path()
    config_file = base_path / config_name

    if not config_file.exists():
        raise FileNotFoundError(f"Missing configuration file: {config_file}")

    config = configparser.ConfigParser()
    try:
        config.read(config_file, encoding='utf-8')
        CONFIG = config
        logging.info(f"Config loaded from {config_file}")
    except Exception as e:
        raise ValueError(f"Error reading config file {config_file}: {e}")

    return config

def get_path(section, key):
    """Return a Path object for a given config value."""
    if CONFIG is None:
        raise RuntimeError("Config not loaded. Call load_config() first.")
    
    value = CONFIG.get(section, key, fallback=None)
    if not value:
        raise ValueError(f"Missing path setting: [{section}] {key}")
    return Path(value)

def get_flag(section, key, default=False):
    """Return a boolean flag from config (case-insensitive)."""
    if CONFIG is None:
        raise RuntimeError("Config not loaded. Call load_config() first.")
    
    val = CONFIG.get(section, key, fallback=str(default)).strip().lower()
    return val in ("true", "1", "yes", "on")

def get_value(section, key, default=None):
    """Return a string value from config."""
    if CONFIG is None:
        raise RuntimeError("Config not loaded. Call load_config() first.")
    
    return CONFIG.get(section, key, fallback=default)

def get_int(section, key, default=0):
    """Return an integer value from config."""
    if CONFIG is None:
        raise RuntimeError("Config not loaded. Call load_config() first.")
    
    try:
        return CONFIG.getint(section, key, fallback=default)
    except ValueError:
        logging.warning(f"Invalid integer for [{section}] {key}, using default: {default}")
        return default