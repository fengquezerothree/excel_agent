import asyncio
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools

def get_first_model_name():
    """获取第一个可用的模型名称"""
    try:
        from openai import OpenAI
        client = OpenAI(
            base_url="http://10.180.116.5:6390/v1",
            api_key="dummy"
        )
        models = client.models.list()
        return models.data[0].id
    except Exception as e:
        print(f"获取模型名称失败: {e}")
        raise

async def test_tool_calling():
    """测试工具调用功能"""
    
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
            model=model_name,
            temperature=0
        )
        
        # 3. 加载 MCP 工具
        async with client.session("excel") as session:
            tools = await load_mcp_tools(session)
            print(f"🔧 加载了 {len(tools)} 个工具")
            
            # 4. 测试工具绑定
            llm_with_tools = llm.bind_tools(tools)
            
            # 5. 发送测试消息
            system_prompt = """你是一个专业的Excel数据分析师。你的任务是：
1. 必须首先使用提供的工具读取Excel文件数据
2. 基于读取的数据进行分析
3. 提供详细的分析报告

重要：你必须先调用工具获取数据，然后再进行分析。不要在没有读取数据的情况下给出答案。

可用工具：
""" + "\n".join([f"- {tool.name}: {tool.description}" for tool in tools])
            
            user_query = "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，请分析用户主要关注哪些问题，并给出一份分析报告。"
            
            messages = [
                HumanMessage(content=system_prompt),
                HumanMessage(content=user_query)
            ]
            
            print("🤖 发送请求给LLM...")
            response = llm_with_tools.invoke(messages)
            
            print(f"🔍 响应类型: {type(response)}")
            print(f"🔍 响应内容: {response.content[:200]}...")
            print(f"🔍 是否有tool_calls属性: {hasattr(response, 'tool_calls')}")
            
            if hasattr(response, 'tool_calls') and response.tool_calls:
                print(f"🔍 工具调用数量: {len(response.tool_calls)}")
                for i, tool_call in enumerate(response.tool_calls):
                    print(f"🔍 工具调用 {i+1}: {tool_call}")
            else:
                print("❌ 没有检测到工具调用")
                
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_tool_calling()) 