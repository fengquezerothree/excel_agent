from agents import set_default_openai_api, set_default_openai_client, Runner, set_tracing_disabled
from openai import AsyncOpenAI
from config_loader import get_model_service_config, get_model_name

# ③ 禁用 OpenAI Agents SDK 的 trace 功能，避免 API 密钥警告
set_tracing_disabled(True)

# ① 先把默认 API 切成 chat_completions
set_default_openai_api("chat_completions")   # ★ 关键修复

# ② 使用配置加载器获取模型服务配置
model_config = get_model_service_config()
client = AsyncOpenAI(
    base_url=model_config["base_url"], 
    api_key=model_config["api_key"]
)
set_default_openai_client(client)

# 使用配置加载器获取模型名称
model = get_model_name()

# advanced_usage.py - 带有 MCP 服务器的用法
import asyncio
from agents import Agent, Runner
from agents.mcp import MCPServerStreamableHttp
from config_loader import get_mcp_server_config

async def advanced_excel_agent():
    """使用 MCP 服务器的高级代理示例"""
    
    # 1. 使用配置加载器设置 MCP 服务器
    mcp_config = get_mcp_server_config()
    mcp_server = MCPServerStreamableHttp(
        params={"url": mcp_config["url"]},
        cache_tools_list=True
    )
    
    try:
        # 2. 连接到 MCP 服务器
        await mcp_server.connect()
        print("已连接到 MCP 服务器")
        
        # 3. 创建带有 MCP 工具的代理
        agent = Agent(
            name="excel-mcp-agent",
            instructions="""
            你是一个专业的 Excel 处理代理。
            使用提供的 MCP 工具来操作 Excel 文件。
            """,
            model=model,
            mcp_servers=[mcp_server]
        )
        
        # 4. 创建客户端和运行器

        runner = Runner()
        
        # 5. 执行任务
        result = await runner.run(
            agent,
            # input="读取 data.xlsx 文件的第一个工作表，显示 A1:C10 的内容"
            input="20250703it.xlsx，读取这个excel的Sheet1的前300行的a到d列，分析一下，用户主要关注些什么问题？给我一个分析报告"
        )
        
        print("=== MCP 代理执行结果 ===")
        print(result.final_output)
        
    except Exception as e:
        print(f"MCP 代理错误: {e}")
    
    finally:
        # 6. 清理资源
        if mcp_server:
            await mcp_server.cleanup()
            print("已关闭 MCP 服务器连接")

if __name__ == "__main__":
    asyncio.run(advanced_excel_agent())
