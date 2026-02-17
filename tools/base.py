import json
import os
from datetime import datetime
import time
from config import WORKSPACE_CONFIG_FILE

def load_workspace_config():
    """Load workspace configuration from file"""
    if os.path.exists(WORKSPACE_CONFIG_FILE):
        try:
            with open(WORKSPACE_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return None
    return None

def get_workspace_path():
    """Get current workspace path"""
    config = load_workspace_config()
    if config and "path" in config:
        return config["path"]
    return None

def get_current_time():
    """Get detailed structured information about the current time"""
    now = datetime.now()
    return {
        "year": now.year,
        "month": now.month,
        "day": now.day,
        "hour": now.hour,
        "minute": now.minute,
        "second": now.second,
        "weekday": now.strftime("%A"),
        "timezone": "UTC+08:00",
        "iso_format": now.isoformat(),
        "readable_format": now.strftime("%Y-%m-%d %H:%M:%S"),
        "day_of_year": now.timetuple().tm_yday,
        "is_daylight_saving": time.localtime().tm_isdst > 0
    }
