# Microsoft Integrations Monolith

A comprehensive monolithic suite for integrating with various Microsoft services and applications. This project provides a unified development experience while maintaining modular architecture for different Microsoft service integrations.

## 🚀 Quick Start

### Prerequisites
- Node.js 16 or higher
- npm or yarn
- Microsoft Office applications (for Office Add-ins)

### Installation
```bash
# Clone the repository
git clone <repository-url>
cd msft-integrations

# Install all dependencies across the monolith
npm run setup
```

### Development Commands
```bash
# Office Integrations
npm run word-webpack    # Start Word Add-in development server
npm run word-start      # Start Word Add-in for sideloading

# AI Services
npm run ai-setup        # Setup AI services (install deps + models)
npm run llm-gateway     # Start LLM Gateway server
npm run local-llm       # Start local model server (Ollama)
npm run ai-status       # Check AI services status

# General
npm run setup           # Install all dependencies across monolith
npm run help            # Show all available commands
npm run install-all     # Install dependencies for all integrations
```

## 📁 Project Structure

```
msft-integrations/
├── integrations/          # Individual integration modules
│   ├── word-addin/       # ✅ Word Office Add-in (Active)
│   ├── excel-addin/      # 🚧 Excel integration (Planned)
│   ├── powerpoint-addin/ # 🚧 PowerPoint integration (Planned)
│   ├── teams-app/        # 🚧 Teams application (Planned)
│   └── outlook-addin/    # 🚧 Outlook add-in (Planned)
├── ai-services/          # 🆕 AI/LLM components (Independent processes)
│   ├── llm-gateway/      # LiteLLM proxy server for unified LLM access
│   ├── local-llm/        # Local model server (Ollama/Llama)
│   └── shared/           # AI utilities and prompt templates
├── shared/               # Shared utilities and common code
│   └── utils/           # Common functions, logging, config, AI client
├── scripts/              # Build and development scripts
│   ├── setup-all.js     # Dependency installation script
│   ├── setup-ai.py      # 🆕 AI services setup automation
│   └── dev.js           # Development command router
└── docs/                # Comprehensive documentation
    ├── ARCHITECTURE.md   # System architecture overview
    ├── DEVELOPMENT.md    # Development guidelines
    └── DEPLOYMENT.md     # Deployment strategies
```

## 🎯 Current Features

### Word Office Add-in ✅
- **Text Operations**: Insert and manipulate text content
- **Selection Handling**: Get and format selected text
- **Table Creation**: Insert formatted tables
- **Document Info**: Retrieve document metadata
- **Error Handling**: Robust error management with logging

### AI Services ✅
- **LLM Gateway**: Unified interface for multiple AI providers (local, OpenAI, Azure)
- **Local Models**: Run Llama, Mistral, and other models locally via Ollama
- **Provider Switching**: Seamlessly switch between local and cloud AI models
- **Office Integration**: AI client library for Office applications
- **Prompt Templates**: Pre-built templates for document analysis, text improvement
- **REST API**: OpenAI-compatible API for easy integration

### Shared Infrastructure ✅
- **Common Utilities**: Logger, Config, Office.js helpers, AI client
- **Development Tools**: Automated setup and development scripts
- **Documentation**: Comprehensive guides and architecture docs
- **Monolithic Benefits**: Shared dependencies and consistent patterns

## Available Integrations

### ✅ Word Integration (Office Add-in)
**Location**: `integrations/word-addin/`

A fully functional Word Office Add-in that demonstrates:
- Text insertion and manipulation
- Document formatting
- Table creation
- Document statistics
- Real-time interaction with Word documents

[📖 See Word Add-in README](integrations/word-addin/README.md)

### 🤖 AI Services Integration
**Location**: `ai-services/`

A comprehensive AI/LLM infrastructure that provides:
- **LLM Gateway**: Unified proxy for multiple AI providers using liteLLM
- **Local Model Server**: Run models like Llama2, Mistral locally via Ollama
- **Provider Flexibility**: Switch between local, OpenAI, and Azure AI seamlessly
- **Office Integration**: AI client library for adding AI features to Office add-ins
- **Document Analysis**: Pre-built prompts for text analysis and improvement
- **REST API**: OpenAI-compatible endpoints for easy integration

[📖 See AI Services README](ai-services/README.md)

### 🚧 Future Integrations

- **Excel Add-in**: Spreadsheet automation and data manipulation
- **PowerPoint Add-in**: Presentation generation and slide management  
- **Teams App**: Custom Teams application with bots and tabs
- **Outlook Add-in**: Email automation and calendar integration
- **SharePoint Integration**: Document management and workflow automation

## Getting Started

1. **Clone the repository**:
   ```bash
   git clone https://github.com/henninggit/msft-integrations.git
   cd msft-integrations
   ```

2. **Choose an integration** and navigate to its directory:
   ```bash
   cd integrations/word-addin
   ```

3. **Follow the specific README** for that integration

## 🛠️ Development

### Architecture Overview
This project uses a **monolithic architecture** with modular design:
- **Single Repository**: All integrations in one place
- **Shared Resources**: Common utilities, configuration, and dependencies
- **Independent Modules**: Each integration can be developed and deployed separately
- **Consistent Patterns**: Unified development experience across all integrations
- **AI-Enhanced**: Local and cloud AI capabilities integrated throughout

### Communication Architecture
- **Office Integrations** ↔ **REST API** ↔ **LLM Gateway** ↔ **AI Providers**
- **Process Independence**: AI services run separately from Office integrations
- **Language Agnostic**: JavaScript Office add-ins communicate with Python AI services
- **Scalable Design**: Easy to move AI services to cloud or scale independently

### Adding New Integrations
1. Create new directory under `integrations/`
2. Follow the established project structure
3. Use shared utilities from `shared/utils/`
4. Integrate with AI services via `AIServiceClient`
5. Add npm scripts to main `package.json`
6. Update development scripts in `scripts/dev.js`

### Best Practices
- Use shared `Logger` for consistent logging
- Leverage `Config` utility for environment management
- Follow `OfficeHelper` patterns for Office.js operations
- Use `AIServiceClient` for AI integrations
- Implement proper error handling throughout

## 📚 Documentation

- **[Architecture Guide](docs/ARCHITECTURE.md)** - System design and component overview
- **[Development Guide](docs/DEVELOPMENT.md)** - Development setup and best practices  
- **[Deployment Guide](docs/DEPLOYMENT.md)** - Production deployment strategies
- **[AI Services README](ai-services/README.md)** - AI infrastructure and setup guide
- **[Word Add-in README](integrations/word-addin/README.md)** - Word integration details
- **[Shared Utils README](shared/README.md)** - Common utilities documentation

## 🔧 Technical Stack

- **Frontend**: HTML, CSS, JavaScript (ES6+)
- **Build Tool**: Webpack with development server
- **Office Integration**: Office.js APIs
- **AI Services**: Python + FastAPI + liteLLM
- **Local AI**: Ollama for running Llama, Mistral, etc.
- **API Gateway**: LiteLLM proxy for unified AI provider access
- **Microsoft Services**: Microsoft Graph API, Azure AD
- **Development**: Node.js, npm scripts, hot reload
- **Architecture**: Monolithic with modular design

## 🎯 Roadmap

### Phase 1: Foundation ✅
- [x] Word Office Add-in with core features
- [x] Shared utilities and common code
- [x] Development tooling and scripts
- [x] Comprehensive documentation
- [x] AI services infrastructure
- [x] LLM Gateway with provider switching
- [x] Local model server integration

### Phase 2: AI Enhancement 🚧
- [ ] AI-powered document analysis in Word add-in
- [ ] Text improvement and writing assistance
- [ ] Smart content generation
- [ ] Document summarization features

### Phase 3: Expansion 📋
- [ ] Excel Office Add-in with AI data analysis
- [ ] PowerPoint Office Add-in with AI presentation generation
- [ ] Microsoft Teams application
- [ ] Outlook Add-in with AI email assistance

### Phase 4: Advanced Features 🔮
- [ ] Cross-application AI workflows
- [ ] Advanced analytics and reporting
- [ ] Custom AI model fine-tuning
- [ ] Enterprise deployment features

## 🚀 Quick Start Guide

### For Office Integration Development:
```bash
npm run setup          # Install all dependencies
npm run word-webpack   # Start Word add-in development
```

### For AI Services Development:
```bash
npm run ai-setup       # Setup AI services + download models
npm run llm-gateway    # Start LLM gateway
npm run ai-status      # Check services status
```

### For Full Stack Development:
```bash
npm run setup          # Install all dependencies
npm run ai-setup       # Setup AI services
npm run llm-gateway    # Terminal 1: Start AI gateway
npm run word-webpack   # Terminal 2: Start Word development
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/new-integration`
3. Follow the established patterns and use shared utilities
4. Test your integration thoroughly
5. Update documentation as needed
6. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

- Check the [Development Guide](docs/DEVELOPMENT.md) for common issues
- Review the [Architecture Guide](docs/ARCHITECTURE.md) for system understanding
- Consult the [AI Services README](ai-services/README.md) for AI-related questions
- Open an issue for bugs or feature requests
- Consult Microsoft's Office Add-ins documentation for platform-specific questions

---

## 🎉 What Makes This Project Special

This Microsoft Integrations Monolith combines the **best of both worlds**:

- **🏢 Enterprise-Ready Office Integrations** - Professional Office add-ins with robust error handling
- **🤖 Cutting-Edge AI Capabilities** - Local and cloud AI models seamlessly integrated
- **🔧 Developer-Friendly Architecture** - Monolithic benefits with modular flexibility
- **📚 Comprehensive Documentation** - Everything you need to extend and deploy
- **🚀 Production-Ready** - From development to deployment, fully covered

Whether you're building simple Office automations or complex AI-powered document workflows, this project provides the foundation to build, scale, and deploy with confidence!