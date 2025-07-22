"""
Shared utilities for AI services
Common functions and classes used across LLM components
"""

import json
import time
from typing import Dict, Any, Optional, List
from dataclasses import dataclass
from enum import Enum

class LLMProvider(Enum):
    LOCAL = "local"
    OPENAI = "openai"
    AZURE = "azure"

@dataclass
class LLMRequest:
    """Standardized LLM request format"""
    prompt: str
    model: str = "gpt-3.5-turbo"
    provider: Optional[LLMProvider] = None
    temperature: float = 0.7
    max_tokens: Optional[int] = None
    stream: bool = False
    context: Optional[Dict[str, Any]] = None

@dataclass
class LLMResponse:
    """Standardized LLM response format"""
    content: str
    model: str
    provider: str
    usage: Optional[Dict[str, int]] = None
    finish_reason: Optional[str] = None
    response_time: Optional[float] = None

class PromptTemplates:
    """Common prompt templates for Office integrations"""
    
    DOCUMENT_ANALYSIS = """
    Analyze the following document content and provide insights:
    
    Content: {content}
    
    Please provide:
    1. A brief summary
    2. Key themes or topics
    3. Suggested improvements
    4. Tone and style assessment
    """
    
    TEXT_IMPROVEMENT = """
    Please improve the following text while maintaining its original meaning:
    
    Original text: {text}
    
    Focus on:
    - Grammar and spelling
    - Clarity and readability
    - Professional tone
    - Conciseness
    """
    
    DATA_INSIGHTS = """
    Analyze the following data and provide insights:
    
    Data: {data}
    
    Please provide:
    1. Key patterns or trends
    2. Notable observations
    3. Potential implications
    4. Recommendations for action
    """
    
    EMAIL_DRAFT = """
    Draft a professional email with the following requirements:
    
    Purpose: {purpose}
    Tone: {tone}
    Key points: {key_points}
    Recipient: {recipient}
    
    Please create a well-structured, professional email.
    """
    
    PRESENTATION_OUTLINE = """
    Create an outline for a presentation on the following topic:
    
    Topic: {topic}
    Audience: {audience}
    Duration: {duration}
    Objective: {objective}
    
    Please provide:
    1. Title suggestions
    2. Main sections/slides
    3. Key points for each section
    4. Conclusion recommendations
    """

class LLMClient:
    """Client for communicating with LLM Gateway"""
    
    def __init__(self, gateway_url: str = "http://localhost:8000"):
        self.gateway_url = gateway_url.rstrip('/')
        self.session = None
    
    async def _make_request(self, endpoint: str, data: Dict[str, Any]) -> Dict[str, Any]:
        """Make HTTP request to gateway"""
        import aiohttp
        
        if not self.session:
            self.session = aiohttp.ClientSession()
        
        url = f"{self.gateway_url}{endpoint}"
        
        async with self.session.post(url, json=data) as response:
            if response.status == 200:
                return await response.json()
            else:
                error_text = await response.text()
                raise Exception(f"LLM Gateway error ({response.status}): {error_text}")
    
    async def chat_completion(self, request: LLMRequest) -> LLMResponse:
        """Send chat completion request"""
        start_time = time.time()
        
        data = {
            "model": request.model,
            "messages": [{"role": "user", "content": request.prompt}],
            "temperature": request.temperature,
            "max_tokens": request.max_tokens,
            "stream": request.stream
        }
        
        if request.provider:
            data["provider"] = request.provider.value
        
        response_data = await self._make_request("/v1/chat/completions", data)
        response_time = time.time() - start_time
        
        # Extract content from OpenAI-format response
        content = response_data["choices"][0]["message"]["content"]
        usage = response_data.get("usage")
        finish_reason = response_data["choices"][0].get("finish_reason")
        
        return LLMResponse(
            content=content,
            model=response_data["model"],
            provider=request.provider.value if request.provider else "unknown",
            usage=usage,
            finish_reason=finish_reason,
            response_time=response_time
        )
    
    async def get_embeddings(self, text: str, model: str = "text-embedding-ada-002") -> List[float]:
        """Get text embeddings"""
        data = {
            "input": text,
            "model": model
        }
        
        response_data = await self._make_request("/v1/embeddings", data)
        return response_data["data"][0]["embedding"]
    
    async def switch_provider(self, provider: LLMProvider) -> bool:
        """Switch the default provider"""
        try:
            await self._make_request("/v1/provider/switch", {"provider": provider.value})
            return True
        except:
            return False
    
    async def close(self):
        """Close the HTTP session"""
        if self.session:
            await self.session.close()

class OfficeIntegrationHelper:
    """Helper class for Office-specific AI integrations"""
    
    def __init__(self, llm_client: LLMClient):
        self.llm_client = llm_client
    
    async def analyze_word_document(self, content: str) -> str:
        """Analyze Word document content"""
        prompt = PromptTemplates.DOCUMENT_ANALYSIS.format(content=content)
        request = LLMRequest(prompt=prompt, model="gpt-3.5-turbo")
        response = await self.llm_client.chat_completion(request)
        return response.content
    
    async def improve_text(self, text: str) -> str:
        """Improve text quality"""
        prompt = PromptTemplates.TEXT_IMPROVEMENT.format(text=text)
        request = LLMRequest(prompt=prompt, model="gpt-3.5-turbo")
        response = await self.llm_client.chat_completion(request)
        return response.content
    
    async def analyze_excel_data(self, data: str) -> str:
        """Analyze Excel data"""
        prompt = PromptTemplates.DATA_INSIGHTS.format(data=data)
        request = LLMRequest(prompt=prompt, model="gpt-3.5-turbo")
        response = await self.llm_client.chat_completion(request)
        return response.content
    
    async def draft_email(self, purpose: str, tone: str = "professional", 
                         key_points: str = "", recipient: str = "colleague") -> str:
        """Draft an email"""
        prompt = PromptTemplates.EMAIL_DRAFT.format(
            purpose=purpose,
            tone=tone,
            key_points=key_points,
            recipient=recipient
        )
        request = LLMRequest(prompt=prompt, model="gpt-3.5-turbo")
        response = await self.llm_client.chat_completion(request)
        return response.content
    
    async def create_presentation_outline(self, topic: str, audience: str = "general", 
                                        duration: str = "10 minutes", 
                                        objective: str = "inform") -> str:
        """Create presentation outline"""
        prompt = PromptTemplates.PRESENTATION_OUTLINE.format(
            topic=topic,
            audience=audience,
            duration=duration,
            objective=objective
        )
        request = LLMRequest(prompt=prompt, model="gpt-3.5-turbo")
        response = await self.llm_client.chat_completion(request)
        return response.content

# Utility functions
def chunk_text(text: str, max_tokens: int = 4000) -> List[str]:
    """Split text into chunks for processing"""
    # Simple word-based chunking (can be improved with tiktoken for exact tokens)
    words = text.split()
    chunks = []
    current_chunk = []
    current_length = 0
    
    # Rough estimate: 1 token ≈ 4 characters
    max_chars = max_tokens * 4
    
    for word in words:
        word_length = len(word) + 1  # +1 for space
        if current_length + word_length > max_chars and current_chunk:
            chunks.append(' '.join(current_chunk))
            current_chunk = [word]
            current_length = word_length
        else:
            current_chunk.append(word)
            current_length += word_length
    
    if current_chunk:
        chunks.append(' '.join(current_chunk))
    
    return chunks

def format_response_for_office(response: str, format_type: str = "plain") -> str:
    """Format LLM response for Office applications"""
    if format_type == "markdown":
        return response
    elif format_type == "html":
        # Convert basic markdown to HTML
        response = response.replace('\n\n', '</p><p>')
        response = response.replace('\n', '<br>')
        response = f"<p>{response}</p>"
        # Handle bold text
        response = response.replace('**', '<strong>').replace('**', '</strong>')
        return response
    elif format_type == "plain":
        # Remove markdown formatting
        response = response.replace('**', '')
        response = response.replace('*', '')
        response = response.replace('#', '')
        return response
    else:
        return response
