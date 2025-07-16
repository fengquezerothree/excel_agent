# main_excel_agent.py
# 这是一个Excel代理，使用LangGraph框架构建，通过MCP客户端与Excel服务交互
import os, asyncio
from datetime import date
from typing import Annotated
from typing_extensions import TypedDict
from openai import OpenAI

from langchain_mcp_adapters.client import MultiServerMCPClient  # MCP 客户端
from langchain_openai import ChatOpenAI
from langchain_core.messages import AnyMessage, HumanMessage
from langgraph.prebuilt import ToolNode, tools_condition
from langgraph.graph import StateGraph, START
from langgraph.graph.message import add_messages


def get_first_model_name():
    """
    获取可用的第一个模型名称
    通过OpenAI客户端连接到服务器并获取可用模型列表
    如果获取失败则抛出异常
    """
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

# 获取模型名称
model_name = get_first_model_name()

# 1. Agent 状态定义
class AgentState(TypedDict):
    """
    定义Agent状态，使用TypedDict
    messages: 消息列表，用于存储对话历史
    """
    messages: Annotated[list[AnyMessage], add_messages]

# 2. 启动 MCP 客户端并加载工具
async def setup_agent(client: MultiServerMCPClient):
    """
    设置Agent环境
    1. 使用已有的MCP客户端
    2. 获取可用工具
    3. 创建ChatOpenAI模型并绑定工具
    返回: 带工具的模型和工具列表
    """
    # 获取工具列表
    tools = await client.get_tools()
    
    # 创建LLM模型并绑定工具
    model = ChatOpenAI(
        base_url="http://10.180.116.5:6390/v1",
        api_key="dummy",
        model=model_name
    )
    model_with_tools = model.bind_tools(tools)
    
    return model_with_tools, tools

# 3. 构建 LangGraph 图
async def build_graph(model_with_tools, tools):
    """
    构建LangGraph执行图
    1. 创建工具节点
    2. 定义模型节点处理函数
    3. 构建状态图，定义节点和边
    返回: 编译后的图
    """
    tool_node = ToolNode(tools)

    async def model_node(state: AgentState):
        """模型节点处理函数，接收状态并返回模型响应"""
        response = await model_with_tools.ainvoke(state["messages"])
        return {"messages": [response]}

    builder = StateGraph(AgentState)
    builder.add_node("model", model_node)
    builder.add_node("tools", tool_node)
    builder.add_edge(START, "model")
    builder.add_conditional_edges("model", tools_condition)
    builder.add_edge("tools", "model")

    graph = builder.compile()
    return graph

# 4. 主运行逻辑
async def main():
    """
    主函数，运行Excel分析任务
    使用 async with 确保 MultiServerMCPClient 自动关闭
    """
    try:
        async with MultiServerMCPClient({
            "excel": {
                "transport": "streamable_http",
                "url": "http://localhost:8007/mcp",
            }
        }) as client:
            # 准备模型和工具
            model_with_tools, tools = await setup_agent(client)
            graph = await build_graph(model_with_tools, tools)
            
            # 用户查询
            input_query = (
                "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
                "请分析用户主要关注哪些问题，并给出一份分析报告。"
            )
            
            # 流式处理事件并输出结果
            async for event in graph.astream_events(
                {"messages": [HumanMessage(content=input_query)]},
                version="v1"
            ):
                kind = event["event"]
                data = event["data"]
                if kind == "on_chat_model_stream":
                    print(data["chunk"].content, end="", flush=True)
                elif kind == "on_tool_call":
                    print(f"\n\n-- 调用工具: {data['tool_call']['name']} --\n")
                elif kind == "on_tool_end":
                    print(f"\n-- 工具结果: {data['output']} --\n")
            
            print("✨ 任务完成！MultiServerMCPClient 已自动关闭。")
    except FileNotFoundError as e:
        print(f"文件未找到: {e}")
    except ConnectionError as e:
        print(f"MCP客户端连接错误: {e}")
    except Exception as e:
        print(f"运行时发生错误: {e}")

if __name__ == "__main__":
    asyncio.run(main())
