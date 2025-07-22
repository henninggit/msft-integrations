# AI Services

This directory contains the AI/LLM components of the Microsoft Integrations project. These services run independently from Office integrations to avoid startup delays and resource conflicts.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                Office Integrations                     │
│            (Word, Excel, PowerPoint, etc.)             │
└─────────────────┬───────────────────────────────────────┘
                  │ REST API calls
                  ▼
┌─────────────────────────────────────────────────────────┐
│                 LLM Gateway                            │
│              (liteLLM proxy)                           │
│    ┌─────────────┬─────────────┬─────────────┐        │
│    │   Local     │   OpenAI    │   Azure     │        │
│    │   Models    │             │   OpenAI    │        │
│    └─────────────┴─────────────┴─────────────┘        │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│                Local LLM Server                        │
│            (Llama, Mistral, etc.)                      │
└─────────────────────────────────────────────────────────┘
```

## Components

### 1. LLM Gateway (`llm-gateway/`)
- **Purpose**: Unified interface for all LLM providers
- **Technology**: Python + liteLLM
- **Features**: 
  - Provider switching (local ↔ cloud)
  - Load balancing
  - Rate limiting
  - Request/response logging
- **API**: REST endpoints compatible with OpenAI format

### 2. Local LLM Server (`local-llm/`)
- **Purpose**: Serve local models (Llama, Mistral, etc.)
- **Technology**: Python + Ollama/llamacpp/vLLM
- **Features**:
  - Model loading/unloading
  - GPU acceleration
  - Model management
- **API**: OpenAI-compatible REST API

### 3. Shared AI Utilities (`shared/`)
- Common prompt templates
- Response parsing utilities
- Model configuration
- Error handling patterns

## Communication Flow

1. **Office Integration** → HTTP request → **LLM Gateway**
2. **LLM Gateway** → Routes to appropriate provider
3. **Provider** → Returns response → **LLM Gateway**
4. **LLM Gateway** → Standardized response → **Office Integration**

## Benefits

- **Independence**: AI services can start/stop without affecting Office integrations
- **Flexibility**: Easy provider switching via configuration
- **Performance**: Office integrations remain lightweight and fast
- **Scalability**: Can deploy AI services separately or move to cloud
- **Development**: Test AI features without loading Office applications

## Getting Started

1. **Start LLM Gateway**: `npm run llm-gateway-start`
2. **Start Local LLM**: `npm run local-llm-start`
3. **Test Integration**: Office add-ins automatically connect to gateway

## Configuration

Environment variables control provider selection:
- `LLM_PROVIDER=local|openai|azure`
- `LOCAL_LLM_URL=http://localhost:11434`
- `OPENAI_API_KEY=your-key`
- `AZURE_OPENAI_ENDPOINT=your-endpoint`
