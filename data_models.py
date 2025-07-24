from typing import Optional, List, Dict, Any, Union
from pydantic import BaseModel, Field
from enum import Enum
from langchain_core.messages import (
    BaseMessage, HumanMessage, AIMessage, ToolMessage, SystemMessage,
    messages_from_dict, messages_to_dict, message_to_dict, convert_to_messages, convert_to_openai_messages
)


class MessageRole(str, Enum):
    """消息角色枚举"""
    SYSTEM = "system"
    USER = "user" 
    ASSISTANT = "assistant"
    TOOL = "tool"


class ToolCall(BaseModel):
    """工具调用信息"""
    id: str = Field(..., description="工具调用ID")
    name: str = Field(..., description="工具名称")
    type: str = Field(default="function", description="工具类型")
    args: Dict[str, Any] = Field(default_factory=dict, description="工具参数")


class TokenUsage(BaseModel):
    """Token使用统计"""
    input_tokens: Optional[int] = None
    output_tokens: Optional[int] = None
    total_tokens: Optional[int] = None


class MessageMetadata(BaseModel):
    """消息元数据"""
    model_name: Optional[str] = None
    finish_reason: Optional[str] = None
    token_usage: Optional[TokenUsage] = None


class ChatMessage(BaseModel):
    """聊天消息模型"""
    role: MessageRole = Field(..., description="消息角色")
    content: Optional[str] = Field(default="", description="消息文本内容")
    
    # Agent特有字段
    tool_calls: Optional[List[ToolCall]] = Field(default=None, description="工具调用列表")
    tool_call_id: Optional[str] = Field(default=None, description="对应的工具调用ID（用于tool角色）")
    
    # 元数据
    metadata: Optional[MessageMetadata] = Field(default=None, description="消息元数据")
    
    # 标识字段
    message_id: Optional[str] = Field(default=None, description="消息唯一标识")


class ChatRequest(BaseModel):
    """聊天请求模型"""
    messages: List[ChatMessage] = Field(..., description="对话历史消息列表")
    
    # 用户和会话信息
    user_id: Optional[str] = Field(default="test", description="用户ID，立讯工号")
    session_id: Optional[str] = Field(default=None, description="会话ID，用于跟踪多轮对话")
    
    # 模型和配置
    model: Optional[str] = Field(default="default", description="指定模型名称")
    stream: Optional[bool] = Field(default=False, description="是否使用流式响应")


class ChatResponse(BaseModel):
    """聊天响应模型"""
    message: ChatMessage = Field(..., description="响应消息")
    
    # 执行状态
    status: str = Field(default="success", description="执行状态")
    error_message: Optional[str] = Field(default=None, description="错误信息")
    
    # Agent执行信息
    iteration_count: Optional[int] = Field(default=None, description="实际迭代次数")
    tools_used: Optional[List[str]] = Field(default=None, description="使用的工具列表")
    execution_time: Optional[float] = Field(default=None, description="总执行时间")
    
    # 元数据
    total_token_usage: Optional[TokenUsage] = Field(default=None, description="总Token使用量")
    model_info: Optional[Dict[str, Any]] = Field(default=None, description="模型信息")


# 消息转换工具函数

def chat_request_to_langchain_messages(chat_request: ChatRequest) -> List[BaseMessage]:
    """
    将 ChatRequest 转换为 LangChain BaseMessage 列表
    
    Args:
        chat_request: 用户请求对象
        
    Returns:
        LangChain BaseMessage 列表
    """
    # 转换为字典格式，符合 OpenAI API 标准
    message_dicts = []
    
    for msg in chat_request.messages:
        msg_dict = {
            "role": msg.role.value,  # MessageRole enum 转为字符串
            "content": msg.content or ""
        }
        
        # 处理工具调用（AI消息）
        if msg.tool_calls:
            msg_dict["tool_calls"] = [
                {
                    "id": tc.id,
                    "type": tc.type,
                    "function": {
                        "name": tc.name,
                        "arguments": tc.args
                    }
                }
                for tc in msg.tool_calls
            ]
        
        # 处理工具消息的 tool_call_id
        if msg.tool_call_id:
            msg_dict["tool_call_id"] = msg.tool_call_id
            
        # 添加消息ID
        if msg.message_id:
            msg_dict["id"] = msg.message_id
            
        message_dicts.append(msg_dict)
    
    # 使用官方转换器，自动处理 role 映射和工具调用解析
    return convert_to_messages(message_dicts)


def langchain_messages_to_chat_response(
    messages: List[BaseMessage], 
    model_name: Optional[str] = None,
    **kwargs
) -> ChatResponse:
    """
    将 LangChain Messages 转换为 ChatResponse
    
    Args:
        messages: LangChain BaseMessage 列表  
        model_name: 模型名称
        **kwargs: 其他响应参数（iteration_count, tools_used, execution_time等）
        
    Returns:
        ChatResponse 对象
    """
    if not messages:
        raise ValueError("Messages list cannot be empty")
    
    # 获取最后一条消息作为响应消息
    last_message = messages[-1]
    
    # 使用官方转换器获取消息字典
    msg_dict = message_to_dict(last_message)
    msg_data = msg_dict["data"]
    
    # 角色映射
    role_mapping = {
        "ai": MessageRole.ASSISTANT,
        "human": MessageRole.USER, 
        "system": MessageRole.SYSTEM,
        "tool": MessageRole.TOOL
    }
    
    # 转换工具调用
    tool_calls = None
    if msg_data.get("tool_calls"):
        tool_calls = [
            ToolCall(
                id=tc["id"],
                name=tc["name"], 
                type=tc.get("type", "function"),
                args=tc["args"]
            )
            for tc in msg_data["tool_calls"]
        ]
    
    # 处理 token 使用情况
    token_usage = None
    if usage_meta := msg_data.get("usage_metadata"):
        token_usage = TokenUsage(
            input_tokens=usage_meta.get("input_tokens"),
            output_tokens=usage_meta.get("output_tokens"), 
            total_tokens=usage_meta.get("total_tokens")
        )
    
    # 构建元数据
    metadata = MessageMetadata(
        model_name=model_name,
        finish_reason=msg_data.get("response_metadata", {}).get("finish_reason"),
        token_usage=token_usage
    )
    
    # 构建响应消息
    response_message = ChatMessage(
        role=role_mapping.get(msg_dict["type"], MessageRole.ASSISTANT),
        content=msg_data.get("content", ""),
        tool_calls=tool_calls,
        tool_call_id=msg_data.get("tool_call_id"), 
        metadata=metadata,
        message_id=msg_data.get("id")
    )
    
    return ChatResponse(
        message=response_message,
        **kwargs
    )


def chat_messages_to_openai_format(chat_messages: List[ChatMessage]) -> List[Dict[str, Any]]:
    """
    将 ChatMessage 列表转换为 OpenAI API 格式
    
    Args:
        chat_messages: ChatMessage 列表
        
    Returns:
        OpenAI API 格式的消息列表
    """
    # 先转为 LangChain Messages
    lc_messages = chat_request_to_langchain_messages(
        ChatRequest(messages=chat_messages)
    )
    
    # 再转为 OpenAI 格式
    return convert_to_openai_messages(lc_messages)


