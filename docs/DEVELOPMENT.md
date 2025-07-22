# Development Guide

## Getting Started

### Initial Setup
```bash
# Clone the repository
git clone <repository-url>
cd msft-integrations

# Install all dependencies across the monolith
npm run setup
```

### Development Workflow

1. **Start development server for specific integration:**
   ```bash
   npm run word-dev     # Word Add-in development
   npm run excel-dev    # Excel Add-in (coming soon)
   npm run teams-dev    # Teams app (coming soon)
   ```

2. **Build for production:**
   ```bash
   npm run word-build   # Build Word Add-in
   ```

3. **Get help:**
   ```bash
   npm run help         # Show all available commands
   ```

## Project Architecture

### Monolithic Structure
This project uses a monolithic architecture where all Microsoft service integrations are housed in a single repository but organized into separate modules.

**Benefits:**
- Shared utilities and common code
- Consistent development experience
- Easier deployment and versioning
- Cross-integration features possible

**Structure:**
```
msft-integrations/
├── integrations/          # Individual integration modules
│   ├── word-addin/       # Word Office Add-in
│   ├── excel-addin/      # Excel integration (planned)
│   ├── powerpoint-addin/ # PowerPoint integration (planned)
│   ├── teams-app/        # Teams application (planned)
│   └── outlook-addin/    # Outlook add-in (planned)
├── shared/               # Shared utilities and code
│   └── utils/           # Common functions
├── scripts/              # Build and development scripts
└── docs/                # Documentation
```

### Adding New Integrations

1. Create new directory under `integrations/`
2. Add package.json with integration-specific dependencies
3. Use shared utilities from `shared/utils/`
4. Add npm scripts to main package.json
5. Update dev.js script with new commands

## Development Best Practices

### Code Organization
- Keep integration-specific code in respective directories
- Use shared utilities for common functionality
- Follow consistent error handling patterns
- Use the shared Logger for consistent logging

### Environment Management
- Use `Config` utility for environment-specific settings
- Keep sensitive data in environment variables
- Use HTTP for development, HTTPS for production

### Testing
- Write tests for shared utilities
- Test each integration independently
- Use consistent testing patterns across integrations

### Office Add-ins Specific

#### SSL Certificate Issues
For development, use HTTP instead of HTTPS to avoid certificate issues:
- Set `useHTTPS: false` in webpack config
- Use `http://localhost:3000` in manifest.xml
- Word/Excel/PowerPoint accept HTTP for development

#### Manifest Files
- Keep manifest files in integration root directory
- Use environment-specific URLs
- Test sideloading process regularly

#### Office.js Best Practices
- Always check Office.js availability
- Use error handling wrappers from OfficeHelper
- Test across different Office versions

## Troubleshooting

### Common Issues

1. **Add-in not loading:**
   - Check manifest.xml URL matches development server
   - Verify Office application supports the add-in type
   - Clear Office cache if needed

2. **SSL Certificate errors:**
   - Switch to HTTP for development
   - Regenerate certificates if using HTTPS

3. **Dependencies not found:**
   - Run `npm run setup` to install all dependencies
   - Check if you're in the correct directory

4. **Port conflicts:**
   - Change port in webpack config and manifest.xml
   - Use `Config.getServerUrl()` for dynamic port handling
