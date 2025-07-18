# excel_agent_with_custom_workflow.py
import asyncio
from typing import TypedDict, List, Dict, Any, Union
from openai import OpenAI
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage, BaseMessage
from langchain_core.tools import BaseTool
from langgraph.graph import StateGraph, END
from langgraph.prebuilt import ToolNode
from pydantic import SecretStr


class AgentState(TypedDict):
    """代理状态定义"""
    messages: List[BaseMessage]
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
    
    def _create_workflow(self):
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
    
    async def _agent_node(self, state: AgentState) -> AgentState:
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
        response = await self.llm.bind_tools(self.tools).ainvoke(messages)
        
        # 打印完整的模型响应
        print("\n" + "="*50)
        print("🧠 模型响应内容:")
        print("="*50)
        print(response.content)
        print("="*50)
        
        # 安全检查tool_calls属性
        tool_calls = getattr(response, 'tool_calls', None)
        if tool_calls:
            print(f"🔧 检测到 {len(tool_calls)} 个工具调用:")
            for i, tool_call in enumerate(tool_calls):
                print(f"  📋 工具 {i+1}: {tool_call.get('name', 'unknown')} - {tool_call.get('args', {})}")
        else:
            print("✅ 模型没有调用工具，准备完成任务")
        
        # 更新状态
        new_state: AgentState = {
            "messages": state["messages"] + [response],
            "iteration_count": state["iteration_count"] + 1,
            "max_iterations": state["max_iterations"]
        }
        
        return new_state
    
    async def _action_node(self, state: AgentState) -> AgentState:
        """执行工具调用"""
        last_message = state["messages"][-1]
        
        # 安全检查tool_calls属性
        tool_calls = getattr(last_message, 'tool_calls', None)
        if tool_calls:
            print(f"\n🛠️ 开始执行 {len(tool_calls)} 个工具调用...")
            
            # 使用 ToolNode 异步执行工具调用
            tool_result = await self.tool_node.ainvoke(state)
            
            # 只打印工具执行的摘要信息
            if isinstance(tool_result, dict) and "messages" in tool_result:
                print(f"✅ 工具执行完成，返回 {len(tool_result['messages'])} 条消息")
                
                # 分析工具返回结果的摘要
                for i, msg in enumerate(tool_result["messages"]):
                    if hasattr(msg, 'content') and msg.content:
                        content_length = len(msg.content)
                        # 如果内容很长，只显示摘要
                        if content_length > 200:
                            print(f"  📄 工具消息 {i+1}: {content_length} 字符 (内容较长，已省略详情)")
                        else:
                            print(f"  📄 工具消息 {i+1}: {msg.content}")
                
                new_state: AgentState = {
                    "messages": tool_result["messages"],
                    "iteration_count": state["iteration_count"],
                    "max_iterations": state["max_iterations"]
                }
                return new_state
            else:
                # 如果工具执行结果格式不对，保持原状态
                print("⚠️ 工具执行结果格式异常，保持原状态")
                return state
        else:
            print("❌ 没有找到工具调用")
            return state
    
    async def _finish_node(self, state: AgentState) -> Dict[str, Any]:
        """完成节点"""
        print("\n🎉 工作流执行完成！")
        
        # 从最后一条AI消息中获取最终答案
        final_answer = "任务已完成"
        if state["messages"]:
            # 从后往前查找最后一条AI消息（不包含工具调用的）
            for message in reversed(state["messages"]):
                if (isinstance(message, AIMessage) and 
                    hasattr(message, 'content') and 
                    message.content and 
                    not getattr(message, 'tool_calls', None)):
                    final_answer = message.content
                    print(f"✅ 成功提取最终分析报告 ({len(final_answer)} 字符)")
                    break
        
        return {"final_answer": final_answer}
    
    def _should_continue(self, state: AgentState) -> str:
        """决定是否继续执行"""
        # 检查迭代次数
        if state["iteration_count"] >= state["max_iterations"]:
            print(f"\n⚠️ 达到最大迭代次数 ({state['max_iterations']})，结束工作流")
            return "finish"
        
        # 检查最后一条消息是否包含工具调用
        if state["messages"]:
            last_message = state["messages"][-1]
            tool_calls = getattr(last_message, 'tool_calls', None)
            if tool_calls:
                print(f"\n🔄 继续下一步：执行工具调用")
                return "continue"
        
        # 如果没有工具调用，则完成
        print(f"\n✅ 模型已完成分析，准备结束工作流")
        return "finish"
    
    async def run(self, query: str, max_iterations: int = 10) -> str:
        """运行工作流"""
        print("🚀 启动 Excel 分析工作流...")
        print(f"📋 用户查询: {query}\n")
        
        # 初始化状态
        initial_state: AgentState = {
            "messages": [HumanMessage(content=query)],
            "iteration_count": 0,
            "max_iterations": max_iterations,
        }
        
        # 运行工作流
        final_state = await self.workflow.ainvoke(initial_state)
        
        # 从最终状态中提取最后的AI分析报告
        final_answer = "工作流执行完成"
        if "messages" in final_state and final_state["messages"]:
            # 从后往前查找最后一条AI消息（不包含工具调用的）
            for message in reversed(final_state["messages"]):
                if (isinstance(message, AIMessage) and 
                    hasattr(message, 'content') and 
                    message.content and 
                    not getattr(message, 'tool_calls', None)):
                    final_answer = message.content
                    print(f"✅ 成功提取最终分析报告 ({len(final_answer)} 字符)")
                    break
        
        return final_answer


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
            api_key=SecretStr("dummy"),
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