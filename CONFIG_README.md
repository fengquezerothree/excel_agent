# Excel MCP Agent 配置系统

## 快速开始

1. **复制配置模板**：
   ```bash
   cp excel_mcp_configs_factory.yaml excel_mcp_configs.yaml
   ```

2. **修改你的配置**：
   编辑 `excel_mcp_configs.yaml` 文件，调整 MCP 服务器地址、模型服务地址等配置。

3. **运行代理**：
   ```bash
   python excel_agent.py
   # 或
   python excel_agent_with_langgraph.py
   # 或
   python excel_agent_with_custom_workflow.py
   ```

## 配置文件说明

### excel_mcp_configs_factory.yaml
- 这是配置模板文件，**不要直接修改**
- 包含所有配置项的默认值和说明
- 作为创建用户配置的参考

### excel_mcp_configs.yaml
- 这是你的个人配置文件
- 从 factory 文件复制后，根据实际环境修改
- 如果不存在，系统会自动使用 factory 文件

## 配置项说明

### MCP 服务器配置
```yaml
mcp_server:
  name: "excel"              # MCP 服务器名称
  transport: "streamable_http"  # 传输方式
  url: "http://your-mcp-server:8007/mcp"  # MCP 服务器地址
```

### 模型服务配置
```yaml
model_service:
  base_url: "http://your-model-server:6390/v1"  # 模型 API 地址
  api_key: "dummy"           # API 密钥
  temperature: 0             # 生成温度
  auto_get_first_model: true # 是否自动获取第一个可用模型
  model_name: ""            # 手动指定模型名称（可选）
```

### 代理配置
```yaml
agent_config:
  max_iterations: 10         # 最大迭代次数
  verbose: true             # 是否启用详细日志
```

## 配置加载器使用

如果你需要在其他代码中使用配置：

```python
from config_loader import (
    get_model_name,
    get_mcp_server_config, 
    get_model_service_config,
    get_agent_config,
    get_mcp_client_config
)

# 获取模型名称
model = get_model_name()

# 获取 MCP 配置
mcp_config = get_mcp_server_config()

# 获取模型服务配置
model_config = get_model_service_config()
``` 