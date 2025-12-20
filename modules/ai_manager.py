import os
import json
import datetime
from openai import OpenAI
from modules.path_manager import get_config_path

# 使用统一路径管理
CONFIG_FILE = get_config_path("ai_config.json")
TOKEN_LOG_FILE = get_config_path("token_usage.json")

class TokenManager:
    @staticmethod
    def log_usage(model_name, tokens):
        if not model_name or tokens <= 0: return
        today = datetime.date.today().isoformat()
        
        data = {}
        if os.path.exists(TOKEN_LOG_FILE):
            try:
                with open(TOKEN_LOG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except: pass
        
        if today not in data: data[today] = {}
        current = data[today].get(model_name, 0)
        data[today][model_name] = current + tokens
        
        with open(TOKEN_LOG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)

    @staticmethod
    def get_today_stats():
        today = datetime.date.today().isoformat()
        if not os.path.exists(TOKEN_LOG_FILE): return {}
        try:
            with open(TOKEN_LOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f).get(today, {})
        except: return {}

class AIManager:
    @staticmethod
    def load_config():
        default_conf = {
            "providers": {}, 
            "roles": {       
                "vision": None,
                "brain": None,
                "enable_think": False
            }
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if "providers" in saved: default_conf["providers"] = saved["providers"]
                    if "roles" in saved: default_conf["roles"].update(saved["roles"])
            except: pass
        return default_conf

    @staticmethod
    def save_config(config_dict):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, indent=2, ensure_ascii=False)

    @staticmethod
    def get_client(role_type):
        conf = AIManager.load_config()
        provider_name = conf["roles"].get(role_type)
        if not provider_name: return None, None
        provider_info = conf["providers"].get(provider_name)
        if not provider_info: return None, None
        try:
            client = OpenAI(api_key=provider_info["key"], base_url=provider_info["url"])
            return client, provider_info["model"]
        except: return None, None