# main_excel_agent_simplified.py
import asyncio
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langgraph.prebuilt import create_react_agent
from config_loader import get_model_service_config, get_model_name, get_mcp_client_config


async def main():
    """主函数：使用 create_react_agent 简化 agent 构建"""
    
    # 1. 使用配置加载器设置 MCP 客户端
    client = MultiServerMCPClient(get_mcp_client_config())
    
    try:
        # 2. 使用配置加载器获取模型配置并初始化 LLM
        model_config = get_model_service_config()
        model_name = get_model_name()
        llm = ChatOpenAI(
            base_url=model_config["base_url"],
            api_key=model_config["api_key"],
            model=model_name,
            temperature=model_config.get("temperature", 0)
        )
        
        # 3. 使用 session 加载 MCP 工具
        async with client.session("excel") as session:
            tools = await load_mcp_tools(session)
            
            # 4. 使用 create_react_agent 创建 agent
            agent = create_react_agent(llm, tools)
            
            # 5. 执行查询
            input_query = (
                "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
                "请分析用户主要关注哪些问题，并给出一份分析报告。"
            )
            
            print("🚀 开始执行 Excel 分析任务...")
            print(f"📋 查询内容: {input_query}\n")
            
            # 6. 流式输出结果
            async for chunk in agent.astream(
                {"messages": [("human", input_query)]},
                stream_mode="values"
            ):
                if "messages" in chunk:
                    last_message = chunk["messages"][-1]
                    if hasattr(last_message, 'content') and last_message.content:
                        print(last_message.content)
                        print("\n" + "="*50 + "\n")
    
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}")
    except ConnectionError as e:
        print(f"❌ MCP 客户端连接错误: {e}")
    except Exception as e:
        print(f"❌ 运行时发生错误: {e}")


if __name__ == "__main__":
    print("📊 Excel Agent 启动中...")
    asyncio.run(main())
