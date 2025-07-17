# main_excel_agent_simplified.py
import asyncio
from openai import OpenAI
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langgraph.prebuilt import create_react_agent


def get_first_model_name():
    """获取第一个可用的模型名称"""
    try:
        client = OpenAI(
            base_url="http://10.180.116.5:6390/v1",
            api_key="dummy"
        )
        models = client.models.list()
        return models.data[0].id
    except Exception as e:
        print(f"获取模型名称失败: {e}")
        raise


async def main():
    """主函数：使用 create_react_agent 简化 agent 构建"""
    
    # 1. 设置 MCP 客户端
    client = MultiServerMCPClient({
        "excel": {
            "transport": "streamable_http",
            "url": "http://localhost:8007/mcp",
        }
    })
    
    try:
        # 2. 获取模型名称并初始化 LLM
        model_name = get_first_model_name()
        llm = ChatOpenAI(
            base_url="http://10.180.116.5:6390/v1",
            api_key="dummy",
            model=model_name
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
