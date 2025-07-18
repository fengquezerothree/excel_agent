# excel_agent_with_custom_workflow.py
import asyncio
from typing import TypedDict, List, Dict, Any
from openai import OpenAI
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage
from langchain_core.tools import BaseTool
from langgraph.graph import StateGraph, END
from langgraph.prebuilt import ToolNode


class AgentState(TypedDict):
    """代理状态定义"""
    messages: List[Any]
    iteration_count: int
    max_iterations: int


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


class ExcelWorkflowAgent:
    """使用 LangGraph 工作流编排的 Excel 代理"""
    
    def __init__(self, llm: ChatOpenAI, tools: List[BaseTool]):
        self.llm = llm
        self.tools = tools
        self.tool_node = ToolNode(tools)
        self.workflow = self._create_workflow()
    
    def _create_workflow(self) -> StateGraph:
        """创建工作流程图"""
        workflow = StateGraph(AgentState)
        
        # 添加节点
        workflow.add_node("agent", self._agent_node)
        workflow.add_node("action", self._action_node)
        workflow.add_node("finish", self._finish_node)
        
        # 设置入口点
        workflow.set_entry_point("agent")
        
        # 添加条件边
        workflow.add_conditional_edges(
            "agent",
            self._should_continue,
            {
                "continue": "action",
                "finish": "finish"
            }
        )
        
        # 添加边
        workflow.add_edge("action", "agent")
        workflow.add_edge("finish", END)
        
        return workflow.compile()
    
    def _agent_node(self, state: AgentState) -> Dict[str, Any]:
        """代理节点：决定下一步行动"""
        print(f"🤖 代理思考中... (第 {state['iteration_count'] + 1} 次迭代)")
        
        # 构建系统消息
        system_prompt = """你是一个专业的Excel数据分析师。你的任务是：
1. 必须首先使用提供的工具读取Excel文件数据
2. 基于读取的数据进行分析
3. 提供详细的分析报告

重要：你必须先调用工具获取数据，然后再进行分析。不要在没有读取数据的情况下给出答案。

可用工具：
""" + "\n".join([f"- {tool.name}: {tool.description}" for tool in self.tools])
        
        # 构建消息列表
        messages = [HumanMessage(content=system_prompt)] + state["messages"]
        
        # 调用LLM
        response = self.llm.bind_tools(self.tools).invoke(messages)
        
        # 添加调试信息
        print(f"🔍 LLM响应类型: {type(response)}")
        print(f"🔍 响应内容: {response.content[:200]}...")
        if hasattr(response, 'tool_calls'):
            print(f"🔍 工具调用数量: {len(response.tool_calls) if response.tool_calls else 0}")
            if response.tool_calls:
                for i, tool_call in enumerate(response.tool_calls):
                    print(f"🔍 工具调用 {i+1}: {tool_call}")
        
        # 更新状态
        new_state = {
            "messages": state["messages"] + [response],
            "iteration_count": state["iteration_count"] + 1
        }
        
        # 检查是否有工具调用（仅用于日志输出）
        if hasattr(response, 'tool_calls') and response.tool_calls:
            print(f"🔧 计划使用工具: {response.tool_calls[0]['name']}")
        else:
            print("✅ 代理给出了回答，准备完成")
        
        return new_state
    
    def _action_node(self, state: AgentState) -> Dict[str, Any]:
        """执行工具调用"""
        last_message = state["messages"][-1]
        
        print(f"🛠️ 执行工具调用，共 {len(last_message.tool_calls)} 个工具")
        
        # 使用 ToolNode 执行工具调用
        tool_result = self.tool_node.invoke(state)
        
        print(f"✅ 工具执行完成")
        
        return tool_result
    
    def _finish_node(self, state: AgentState) -> Dict[str, Any]:
        """完成节点"""
        print("🎉 任务完成！")
        
        # 从最后一条AI消息中获取最终答案
        final_answer = "任务已完成"
        if state["messages"]:
            last_message = state["messages"][-1]
            if hasattr(last_message, 'content') and last_message.content:
                final_answer = last_message.content
        
        return {"final_answer": final_answer}
    
    def _should_continue(self, state: AgentState) -> str:
        """决定是否继续执行"""
        # 检查迭代次数
        if state["iteration_count"] >= state["max_iterations"]:
            print(f"⚠️ 达到最大迭代次数 ({state['max_iterations']})")
            return "finish"
        
        # 检查最后一条消息是否包含工具调用
        if state["messages"]:
            last_message = state["messages"][-1]
            if hasattr(last_message, 'tool_calls') and last_message.tool_calls:
                print(f"🔍 检测到工具调用，继续执行")
                return "continue"
        
        # 如果没有工具调用，则完成
        print(f"🔍 没有工具调用，准备完成")
        return "finish"
    
    async def run(self, query: str, max_iterations: int = 10) -> str:
        """运行工作流"""
        print("🚀 启动 Excel 分析工作流...")
        print(f"📋 用户查询: {query}\n")
        
        # 初始化状态
        initial_state = {
            "messages": [HumanMessage(content=query)],
            "iteration_count": 0,
            "max_iterations": max_iterations,
        }
        
        # 运行工作流
        final_state = await self.workflow.ainvoke(initial_state)
        
        return final_state.get("final_answer", "工作流执行完成")


async def main():
    """主函数：使用自定义工作流的 Excel 代理"""
    
    # 1. 设置 MCP 客户端
    client = MultiServerMCPClient({
        "excel": {
            "transport": "streamable_http",
            "url": "http://10.180.39.254:8007/mcp",
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
        
        # 3. 使用 session 加载 MCP 工具
        async with client.session("excel") as session:
            tools = await load_mcp_tools(session)
            print(f"🔧 加载了 {len(tools)} 个工具: {[tool.name for tool in tools]}")
            
            # 4. 创建自定义工作流代理
            agent = ExcelWorkflowAgent(llm, tools)
            
            # 5. 执行查询
            input_query = (
                "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
                "请分析用户主要关注哪些问题，并给出一份分析报告。"
            )
            
            # 6. 运行工作流并获取结果
            result = await agent.run(input_query)
            
            print("\n" + "="*60)
            print("📊 最终分析报告:")
            print("="*60)
            print(result)
    
    except FileNotFoundError as e:
        print(f"❌ 文件未找到: {e}")
    except ConnectionError as e:
        print(f"❌ MCP 客户端连接错误: {e}")
    except Exception as e:
        print(f"❌ 运行时发生错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("📊 自定义工作流 Excel Agent 启动中...")
    asyncio.run(main()) 