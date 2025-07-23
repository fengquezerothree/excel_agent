# excel_agent_with_custom_workflow.py
import asyncio
from typing import TypedDict, List, Dict, Any, Union, Annotated
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain_mcp_adapters.tools import load_mcp_tools
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage, ToolMessage, BaseMessage, SystemMessage
from langchain_core.tools import BaseTool
from langgraph.graph import StateGraph, END, add_messages
from langgraph.prebuilt import ToolNode
from pydantic import SecretStr
from pydantic.type_adapter import P
from config_loader import get_model_service_config, get_model_name, get_mcp_client_config, get_agent_config


class AgentState(TypedDict):
    """代理状态定义"""
    messages: Annotated[List[BaseMessage], add_messages]
    iteration_count: int
    max_iterations: int


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
        system_prompt = """你是一个专业的Excel数据分析师和自动化专家。作为一个高级数据分析代理，你具备以下核心能力和职责：

## 核心身份与职责
你是用户的专业数据分析伙伴，专精于Excel文件的深度分析、数据洞察挖掘和业务价值发现。你的使命是通过精确的数据处理和深入的分析，为用户提供可操作的商业洞察。

## 工作原则与方法论

### 数据驱动决策原则
- 必须基于实际数据进行分析，绝不基于假设或推测
- 先获取真实数据，再进行分析和结论
- 每个结论都要有数据支撑，标明数据来源和分析依据

### 工具调用策略
- 主动使用工具获取所需数据和信息
- 按逻辑顺序执行工具调用：数据读取 → 数据处理 → 分析计算 → 结果验证
- 当单个工具无法完成任务时，智能组合多个工具
- 遇到工具调用失败时，尝试替代方法或调整参数
- 对于复杂任务，将其分解为多个步骤逐步完成

### 分析深度要求
- 不仅要描述数据表面现象，更要挖掘深层趋势和模式
- 识别异常值、数据质量问题和潜在的业务风险
- 提供前瞻性的洞察和建议
- 考虑业务上下文，让分析结果具有实际指导意义

## 专业分析框架

### 专业表达标准
- 使用准确的数据分析术语
- 避免过于技术化的表述，确保业务人员能理解
- 结论要明确、具体、可操作
- 避免冗余信息，聚焦关键洞察


## 交互与协作规范

### 主动性原则
- 主动执行必要的数据获取和分析任务
- 发现重要问题时主动深入挖掘
- 不等待用户明确指示就开始基础数据探索
- 完成主要任务后，主动提供延伸分析建议

### 沟通效率
- 避免不必要的确认和客套话
- 直接开始执行用户请求的分析任务
- 用数据和事实说话，减少主观描述
- 重要发现优先展示，细节信息按需提供

### 问题处理
- 遇到模糊请求时，基于常见业务场景进行合理推断
- 无法获取关键数据时，说明限制并提供替代方案
- 发现数据异常时，及时指出并分析可能原因
- 分析结果与预期不符时，提供可能的解释


"""
        # 历史消息长度
        print(f"历史消息长度(不包含系统消息)：{len(state['messages'])}")

        # 构建消息列表
        messages = [SystemMessage(content=system_prompt)] + state["messages"]
        
        # 调用LLM
        response = await self.llm.bind_tools(self.tools).ainvoke(messages)

        print("\n┌" + "─"*60 + "┐")
        print("│" + " "*18 + "📋 模型完整响应" + " "*18 + "│")
        print("└" + "─"*60 + "┘")
        print(response)

        # 打印完整的模型响应
        print("\n╔" + "═"*48 + "╗")
        print("║" + " "*12 + "🧠 模型响应内容分析" + " "*12 + "║")
        print("╚" + "═"*48 + "╝")
        
        # 检查是否有工具调用
        tool_calls = getattr(response, 'tool_calls', None)
        if tool_calls:
            print("├─ 🔧 模型决定调用工具:")
            for i, tool_call in enumerate(tool_calls):
                tool_name = tool_call.get('name', 'unknown')
                tool_args = tool_call.get('args', {})
                tool_id = tool_call.get('id', 'no-id')
                print(f"│  {i+1}. 工具名称: {tool_name}")
                print(f"│     工具参数: {tool_args}")
                print(f"│     调用ID: {tool_id}")
        elif response.content:
            print("├─ 💬 模型文本响应:")
            print("│  " + response.content.replace('\n', '\n│  '))
        else:
            print("├─ ⚠️ 模型响应为空（无内容且无工具调用）")
        

        
        # 检查是否需要继续迭代
        if tool_calls:
            print(f"└─ 🔄 将执行 {len(tool_calls)} 个工具调用")
        else:
            print("└─ ✅ 模型没有调用工具，准备完成任务")

        # 更新状态 - 只返回新消息，框架会自动追加历史消息
        new_state: AgentState = {
            "messages": [response],
            "iteration_count": state["iteration_count"] + 1,
            "max_iterations": state["max_iterations"]
        }
        
        return new_state
    
    async def _action_node(self, state: AgentState) -> AgentState:
        """执行工具调用"""
        # 打印历史消息条数
        print("\n" + "▼"*30 + " 工具执行区域 " + "▼"*30)
        print(f"📊 当前历史消息数量: {len(state['messages'])}")
        print("─"*75)

        last_message = state["messages"][-1]
        
        # 检查工具调用
        tool_calls = getattr(last_message, 'tool_calls', None)
        if tool_calls:
            print(f"\n🛠️ 开始执行 {len(tool_calls)} 个工具调用...")
            for i, tool_call in enumerate(tool_calls):
                tool_name = tool_call.get('name', 'unknown')
                print(f"  📋 执行工具 {i+1}: {tool_name}")
            
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
                
                # 只返回工具执行产生的新消息，框架会自动追加历史消息
                new_state: AgentState = {
                    "messages": tool_result["messages"],
                    "iteration_count": state["iteration_count"],
                    "max_iterations": state["max_iterations"]
                }
                
                print("▲"*30 + " 工具执行完成 " + "▲"*30)
                return new_state
            else:
                # 如果工具执行结果格式不对，保持原状态
                print("⚠️ 工具执行结果格式异常，保持原状态")
                print("▲"*30 + " 工具执行异常 " + "▲"*30)
                return state
        else:
            print("❌ 没有找到工具调用")
            print("▲"*30 + " 无工具调用 " + "▲"*30)
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
        print("\n" + "◆"*25 + " 流程决策点 " + "◆"*25)
        print("🔍 Agent决定是否继续执行...")
        
        # 检查迭代次数
        if state["iteration_count"] >= state["max_iterations"]:
            print(f"⚠️ 达到最大迭代次数 ({state['max_iterations']})，结束工作流")
            print("◆"*60)
            return "finish"
        
        # 检查最后一条消息是否包含工具调用
        if state["messages"]:
            last_message = state["messages"][-1]
            tool_calls = getattr(last_message, 'tool_calls', None)
            if tool_calls:
                print(f"🔄 继续下一步：执行 {len(tool_calls)} 个工具调用")
                print("◆"*60)
                return "continue"
        
        # 如果没有工具调用，则完成
        print(f"✅ 模型已完成分析，准备结束工作流")
        print("◆"*60)
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
    
    # 1. 使用配置加载器设置 MCP 客户端
    client = MultiServerMCPClient(get_mcp_client_config())
    
    try:
        # 2. 使用配置加载器获取模型配置并初始化 LLM
        # 使用默认模型 qwen2.5-32B
        model_name = "qwen2.5-32B"
        model_config = get_model_service_config(model_name)
        model_name = get_model_name(model_name)
        llm = ChatOpenAI(
            base_url=model_config["base_url"],
            api_key=SecretStr(model_config["api_key"]),
            model=model_name,
            temperature=model_config.get("temperature", 0)
        )
        
        # 3. 使用 session 加载 MCP 工具
        async with client.session("excel") as session:
            tools = await load_mcp_tools(session)
            print(f"🔧 从Ecel MCP server加载了 {len(tools)} 个工具: {[tool.name for tool in tools]}")
            
            # 4. 创建自定义工作流代理
            agent = ExcelWorkflowAgent(llm, tools)
            
            # 5. 执行查询
            input_query = (
                "读取 20250703it.xlsx 的 Sheet1，前300行 A 到 D 列，"
                "请分析用户主要关注哪些问题，并给出一份统计分析报告。"
            )
            
            # 6. 使用配置中的参数运行工作流并获取结果
            agent_cfg = get_agent_config()
            result = await agent.run(input_query, max_iterations=agent_cfg.get("max_iterations", 10))
            
            print("\n" + "★"*20 + " 最终回答 " + "★"*20)
            print(result)
            print("★"*60)
    
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