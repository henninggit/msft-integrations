#!/usr/bin/env python3
"""
LLM Gateway Server
A liteLLM-powered proxy that provides unified access to multiple LLM providers
"""

import os
import asyncio
from typing import Optional, Dict, Any
from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import uvicorn
from dotenv import load_dotenv
from loguru import logger
import litellm
from litellm import completion, acompletion, embedding
from pydantic import BaseModel
import json

# Load environment variables
load_dotenv()

# Configure logging
logger.add("logs/gateway.log", rotation="500 MB", retention="10 days")

# Initialize FastAPI app
app = FastAPI(
    title="LLM Gateway",
    description="Unified interface for multiple LLM providers using liteLLM",
    version="1.0.0"
)

# Configure CORS for Office integrations
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",  # Word add-in dev server
        "https://localhost:3000", # Word add-in production
        "https://*.officeapps.live.com",  # Office Online
        "https://*.office.com",   # Office applications
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configuration
class GatewayConfig:
    def __init__(self):
        self.default_provider = os.getenv("LLM_PROVIDER", "local")
        self.local_llm_url = os.getenv("LOCAL_LLM_URL", "http://localhost:11434")
        self.openai_api_key = os.getenv("OPENAI_API_KEY")
        self.azure_openai_endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
        self.azure_openai_key = os.getenv("AZURE_OPENAI_KEY")
        
        # Configure liteLLM
        if self.openai_api_key:
            litellm.openai_key = self.openai_api_key
        
        if self.azure_openai_endpoint and self.azure_openai_key:
            litellm.azure_key = self.azure_openai_key
            litellm.azure_api_base = self.azure_openai_endpoint

config = GatewayConfig()

# Request/Response Models
class ChatMessage(BaseModel):
    role: str
    content: str

class ChatCompletionRequest(BaseModel):
    model: str = "gpt-3.5-turbo"
    messages: list[ChatMessage]
    temperature: Optional[float] = 0.7
    max_tokens: Optional[int] = None
    stream: Optional[bool] = False
    provider: Optional[str] = None

class EmbeddingRequest(BaseModel):
    input: str | list[str]
    model: str = "text-embedding-ada-002"
    provider: Optional[str] = None

# Provider mapping
def get_model_name(provider: str, model: str) -> str:
    """Map generic model names to provider-specific names"""
    provider_models = {
        "local": {
            "gpt-3.5-turbo": "llama2",
            "gpt-4": "llama2:70b",
            "text-embedding-ada-002": "nomic-embed-text"
        },
        "openai": {
            "gpt-3.5-turbo": "gpt-3.5-turbo",
            "gpt-4": "gpt-4",
            "text-embedding-ada-002": "text-embedding-ada-002"
        },
        "azure": {
            "gpt-3.5-turbo": "azure/gpt-35-turbo",
            "gpt-4": "azure/gpt-4",
            "text-embedding-ada-002": "azure/text-embedding-ada-002"
        }
    }
    
    return provider_models.get(provider, {}).get(model, model)

@app.get("/")
async def health_check():
    """Health check endpoint"""
    return {
        "status": "healthy",
        "service": "LLM Gateway",
        "version": "1.0.0",
        "providers": {
            "local": config.local_llm_url,
            "openai": "configured" if config.openai_api_key else "not configured",
            "azure": "configured" if config.azure_openai_endpoint else "not configured"
        }
    }

@app.post("/v1/chat/completions")
async def chat_completions(request: ChatCompletionRequest):
    """OpenAI-compatible chat completions endpoint"""
    try:
        # Determine provider
        provider = request.provider or config.default_provider
        model_name = get_model_name(provider, request.model)
        
        # Add provider prefix for local models
        if provider == "local":
            model_name = f"ollama/{model_name}"
            litellm.api_base = config.local_llm_url
        
        logger.info(f"Chat completion request: provider={provider}, model={model_name}")
        
        # Prepare messages
        messages = [{"role": msg.role, "content": msg.content} for msg in request.messages]
        
        # Make completion request
        if request.stream:
            response = await acompletion(
                model=model_name,
                messages=messages,
                temperature=request.temperature,
                max_tokens=request.max_tokens,
                stream=True
            )
            
            async def generate():
                async for chunk in response:
                    yield f"data: {json.dumps(chunk)}\n\n"
                yield "data: [DONE]\n\n"
            
            return StreamingResponse(generate(), media_type="text/plain")
        else:
            response = await acompletion(
                model=model_name,
                messages=messages,
                temperature=request.temperature,
                max_tokens=request.max_tokens
            )
            
            return response
            
    except Exception as e:
        logger.error(f"Chat completion error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/v1/embeddings")
async def create_embeddings(request: EmbeddingRequest):
    """OpenAI-compatible embeddings endpoint"""
    try:
        # Determine provider
        provider = request.provider or config.default_provider
        model_name = get_model_name(provider, request.model)
        
        if provider == "local":
            model_name = f"ollama/{model_name}"
            litellm.api_base = config.local_llm_url
        
        logger.info(f"Embedding request: provider={provider}, model={model_name}")
        
        # Make embedding request
        response = await embedding(
            model=model_name,
            input=request.input
        )
        
        return response
        
    except Exception as e:
        logger.error(f"Embedding error: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/v1/models")
async def list_models():
    """List available models"""
    return {
        "object": "list",
        "data": [
            {
                "id": "gpt-3.5-turbo",
                "object": "model",
                "owned_by": "gateway",
                "providers": ["local", "openai", "azure"]
            },
            {
                "id": "gpt-4",
                "object": "model", 
                "owned_by": "gateway",
                "providers": ["local", "openai", "azure"]
            },
            {
                "id": "text-embedding-ada-002",
                "object": "model",
                "owned_by": "gateway", 
                "providers": ["local", "openai", "azure"]
            }
        ]
    }

@app.post("/v1/provider/switch")
async def switch_provider(provider: str):
    """Switch default provider"""
    if provider not in ["local", "openai", "azure"]:
        raise HTTPException(status_code=400, detail="Invalid provider")
    
    config.default_provider = provider
    logger.info(f"Switched default provider to: {provider}")
    
    return {"status": "success", "provider": provider}

if __name__ == "__main__":
    # Create logs directory
    os.makedirs("logs", exist_ok=True)
    
    # Start server
    port = int(os.getenv("GATEWAY_PORT", 8000))
    logger.info(f"Starting LLM Gateway on port {port}")
    
    uvicorn.run(
        "server:app",
        host="127.0.0.1",
        port=port,
        reload=True if os.getenv("DEBUG") else False,
        access_log=True
    )
