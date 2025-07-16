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
async def setup_agent():
    """
    设置Agent环境
    1. 初始化MCP客户端，连接到Excel服务
    2. 获取可用工具
    3. 创建ChatOpenAI模型并绑定工具
    返回: 带工具的模型、工具列表、MCP客户端
    """
    try:
        # 初始化MCP客户端
        client = MultiServerMCPClient({
            "excel": {
                "transport": "streamable_http",
                "url": "http://localhost:8007/mcp",
            }
        })
        # 获取工具列表
        tools = await client.get_tools()
        
        # 创建LLM模型并绑定工具
        model = ChatOpenAI(
                base_url="http://10.180.116.5:6390/v1",
                api_key="dummy",
                model=model_name
            )
        model_with_tools = model.bind_tools(tools)
        
        return model_with_tools, tools, client
    except Exception as e:
        print(f"设置MCP客户端连接失败: {e}")
        raise

# 3. 构建 LangGraph 图
async def build_graph():
    """
    构建LangGraph执行图
    1. 设置Agent，获取模型和工具
    2. 创建工具节点
    3. 定义模型节点处理函数
    4. 构建状态图，定义节点和边
    返回: 编译后的图和MCP客户端
    """
    try:
        # 设置Agent，获取模型和工具
        model_with_tools, tools, client = await setup_agent()
        # 创建工具节点
        tool_node = ToolNode(tools)

        # 定义模型节点处理函数
        async def model_node(state: AgentState):
            """模型节点处理函数，接收状态并返回模型响应"""
            response = await model_with_tools.ainvoke(state["messages"])
            return {"messages": [response]}

        # 构建状态图
        builder = StateGraph(AgentState)
        builder.add_node("model", model_node)  # 添加模型节点
        builder.add_node("tools", tool_node)   # 添加工具节点
        builder.add_edge(START, "model")       # 开始->模型
        builder.add_conditional_edges("model", tools_condition)  # 模型->条件路由
        builder.add_edge("tools", "model")     # 工具->模型

        # 编译图
        graph = builder.compile()
        return graph, client
    except Exception as e:
        print(f"构建LangGraph图失败: {e}")
        if 'client' in locals():
            await client.close()
        raise

# 4. 主运行逻辑
async def main():
    """
    主函数，运行Excel分析任务
    1. 构建执行图
    2. 提交用户查询
    3. 流式处理事件并输出结果
    4. 处理异常并确保资源释放
    """
    client = None
    try:
        # 构建执行图
        graph, client = await build_graph()
        
        # 定义用户查询
        input_query = (
            "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
            "请分析用户主要关注哪些问题，并给出一份分析报告。"
        )
        
        # 流式处理执行图事件
        async for event in graph.astream_events(
            {"messages": [HumanMessage(content=input_query)]}, version="v1"
        ):
            kind = event["event"]
            data = event["data"]
            # 处理模型输出流
            if kind == "on_chat_model_stream":
                chunk = data["chunk"].content
                print(chunk, end="", flush=True)
            # 处理工具调用
            elif kind == "on_tool_call":
                print(f"\n\n-- 调用工具: {data['tool_call']['name']} --\n")
            # 处理工具执行结果
            elif kind == "on_tool_end":
                print(f"\n-- 工具结果: {data['output']} --\n")

    except FileNotFoundError as e:
        # 处理文件不存在的异常
        print(f"文件未找到: {e}")
    except ConnectionError as e:
        # 处理连接错误的异常
        print(f"MCP客户端连接错误: {e}")
    except Exception as e:
        # 处理其他未预期的异常
        print(f"运行时发生错误: {e}")
    finally:
        # 确保MCP客户端正确关闭
        if client:
            try:
                await client.close()
                print("已关闭 MCP 客户端连接")
            except Exception as e:
                print(f"关闭MCP客户端连接时发生错误: {e}")

if __name__ == "__main__":
    # 运行主函数
    asyncio.run(main())
