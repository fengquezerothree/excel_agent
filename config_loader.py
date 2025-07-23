"""
Excel MCP Agent 配置加载器
简洁的配置管理，基于工厂模式设计思想
"""
import yaml
import os
from typing import Dict, Any
from openai import OpenAI


class ConfigLoader:
    """配置加载器 - 使用工厂模式思想"""
    
    def __init__(self, config_file: str = "excel_mcp_configs.yaml"):
        """
        初始化配置加载器
        
        Args:
            config_file: 用户配置文件路径，默认为 excel_mcp_configs.yaml
        """
        self.config_file = config_file
        self._config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if not os.path.exists(self.config_file):
            raise FileNotFoundError(f"配置文件 {self.config_file} 不存在，请创建该文件并配置相关参数")
        
        with open(self.config_file, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    

    
    @property
    def mcp_server_config(self) -> Dict[str, Any]:
        """获取 MCP 服务器配置"""
        return self._config["mcp_server"]
    
    def get_model_service_config(self, model: str) -> Dict[str, Any]:
        """获取指定模型的服务配置"""
        model_service = self._config["model_service"]
        if model not in model_service:
            raise KeyError(f"模型 '{model}' 不存在于配置中，可用模型: {list(model_service.keys())}")
        return model_service[model]
    
    @property
    def agent_config(self) -> Dict[str, Any]:
        """获取代理配置"""
        return self._config["agent_config"]
    
    def get_model_name(self, model: str) -> str:
        """获取指定模型的名称"""
        model_config = self.get_model_service_config(model)
        
        # 如果配置中指定了模型名称，直接返回
        if model_config.get("model_name"):
            return model_config["model_name"]
        
        # 如果启用自动获取第一个模型
        if model_config.get("auto_get_first_model", True):
            try:
                client = OpenAI(
                    base_url=model_config["base_url"],
                    api_key=model_config["api_key"]
                )
                models = client.models.list()
                if models.data:
                    return models.data[0].id
                else:
                    raise Exception("没有可用的模型")
            except Exception as e:
                print(f"❌ 获取模型名称失败: {e}")
                raise
        
        raise Exception("无法获取模型名称，请检查配置")
    
    def get_mcp_client_config(self) -> Dict[str, Dict[str, Any]]:
        """获取 MCP 客户端配置格式"""
        mcp_config = self.mcp_server_config
        return {
            mcp_config["name"]: {
                "transport": mcp_config["transport"],
                "url": mcp_config["url"]
            }
        }


# 全局配置实例（单例模式）
_config_loader = None


def get_config_loader() -> ConfigLoader:
    """获取全局配置加载器实例"""
    global _config_loader
    if _config_loader is None:
        _config_loader = ConfigLoader()
    return _config_loader


# 便捷函数
def get_mcp_server_config() -> Dict[str, Any]:
    """获取 MCP 服务器配置"""
    return get_config_loader().mcp_server_config


def get_model_service_config(model: str) -> Dict[str, Any]:
    """获取指定模型的服务配置"""
    return get_config_loader().get_model_service_config(model)


def get_agent_config() -> Dict[str, Any]:
    """获取代理配置"""
    return get_config_loader().agent_config


def get_model_name(model: str) -> str:
    """获取指定模型的名称"""
    return get_config_loader().get_model_name(model)


def get_mcp_client_config() -> Dict[str, Dict[str, Any]]:
    """获取 MCP 客户端配置"""
    return get_config_loader().get_mcp_client_config() 