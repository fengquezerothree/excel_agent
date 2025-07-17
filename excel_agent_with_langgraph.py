# main_excel_agent.py
import os
import asyncio
from datetime import date
from typing import Annotated
from typing_extensions import TypedDict
from openai import OpenAI

from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langchain_core.messages import AnyMessage, HumanMessage
from langgraph.prebuilt import ToolNode, tools_condition
from langgraph.graph import StateGraph, START
from langgraph.graph.message import add_messages


def get_first_model_name():
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


model_name = get_first_model_name()


class AgentState(TypedDict):
    messages: Annotated[list[AnyMessage], add_messages]


async def setup_agent_with_tools():
    """
    设置 MCP 客户端和 Excel 工具（仅初始化，但不全局 get_tools）
    使用 .session(...) 加载工具
    """
    client = MultiServerMCPClient({
        "excel": {
            "transport": "streamable_http",
            "url": "http://localhost:8007/mcp",
        }
    })

    # 使用 session 获取 MCP 工具
    async with client.session("excel") as session:
        tools = await load_mcp_tools(session)

    # 绑定模型
    model = ChatOpenAI(
        base_url="http://10.180.116.5:6390/v1",
        api_key="dummy",
        model=model_name
    )
    model_with_tools = model.bind_tools(tools)

    return client, model_with_tools, tools


async def build_graph(model_with_tools, tools):
    tool_node = ToolNode(tools)

    async def model_node(state: AgentState):
        response = await model_with_tools.ainvoke(state["messages"])
        return {"messages": [response]}

    builder = StateGraph(AgentState)
    builder.add_node("model", model_node)
    builder.add_node("tools", tool_node)
    builder.add_edge(START, "model")
    builder.add_conditional_edges("model", tools_condition)
    builder.add_edge("tools", "model")
    return builder.compile()


async def main():
    client = None
    try:
        client, model_with_tools, tools = await setup_agent_with_tools()
        graph = await build_graph(model_with_tools, tools)

        input_query = (
            "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
            "请分析用户主要关注哪些问题，并给出一份分析报告。"
        )

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

    except FileNotFoundError as e:
        print(f"文件未找到: {e}")
    except ConnectionError as e:
        print(f"MCP 客户端连接错误: {e}")
    except Exception as e:
        print(f"运行时发生错误: {e}")
    finally:
        if client:
            try:
                await client.close()  # ✅ 正确关闭 MCP 客户端连接
                print("已关闭 MCP 客户端连接")
            except Exception as e:
                print(f"关闭 MCP 客户端连接时出错: {e}")


if __name__ == "__main__":
    asyncio.run(main())
