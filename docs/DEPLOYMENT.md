# Deployment Guide

## Overview

This guide covers deployment strategies for the Microsoft Integrations Monolith across different environments and Microsoft services.

## Office Add-ins Deployment

### Development Deployment (Sideloading)

1. **Build the add-in:**
   ```bash
   npm run word-build
   ```

2. **Start the development server:**
   ```bash
   npm run word-start
   ```

3. **Sideload in Office application:**
   - Word: Insert > Get Add-ins > Upload My Add-in
   - Select the manifest.xml file
   - Add-in appears in ribbon

### Production Deployment Options

#### Option 1: Microsoft AppSource
- Package add-in for AppSource submission
- Requires compliance validation
- Reaches all Office users globally

#### Option 2: Organization Deployment
- Deploy via Microsoft 365 admin center
- Available to organization users only
- Faster deployment process

#### Option 3: SharePoint App Catalog
- Upload to organization's app catalog
- Deploy to specific SharePoint sites
- Good for team-specific add-ins

### Deployment Checklist

- [ ] Update manifest.xml with production URLs
- [ ] Configure HTTPS certificates
- [ ] Test in all target Office applications
- [ ] Validate compliance requirements
- [ ] Set up monitoring and logging
- [ ] Configure error reporting

## Teams Apps Deployment

### Development
```bash
npm run teams-dev    # When available
```

### Production Options
- Microsoft Teams App Store
- Organization app catalog
- Direct upload to Teams

## Shared Infrastructure

### Web Server Requirements
- Node.js 16+ runtime
- HTTPS with valid SSL certificates
- CORS configuration for Office applications
- Static file serving for add-in assets

### Recommended Hosting Platforms
- **Azure App Service** - Integrated Microsoft ecosystem
- **Azure Static Web Apps** - For static content
- **Vercel** - Simple deployment for frontend
- **Heroku** - Quick setup for development

### Environment Variables
```bash
NODE_ENV=production
SERVER_URL=https://your-domain.com
PORT=443
OFFICE_ADDIN_BASE_URL=https://your-domain.com
```

## CI/CD Pipeline

### GitHub Actions Example
```yaml
name: Deploy Office Add-ins

on:
  push:
    branches: [main]

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v2
      
      - name: Setup Node.js
        uses: actions/setup-node@v2
        with:
          node-version: '16'
          
      - name: Install dependencies
        run: npm run setup
        
      - name: Build Word Add-in
        run: npm run word-build
        
      - name: Deploy to Azure
        run: |
          # Your deployment commands here
```

### Azure DevOps Pipeline
```yaml
trigger:
- main

pool:
  vmImage: 'ubuntu-latest'

steps:
- task: NodeTool@0
  inputs:
    versionSpec: '16.x'
    
- script: npm run setup
  displayName: 'Install dependencies'
  
- script: npm run word-build
  displayName: 'Build Word Add-in'
  
- task: AzureWebApp@1
  inputs:
    azureSubscription: 'Your-Azure-Subscription'
    appName: 'your-app-name'
    package: '.'
```

## Security Considerations

### Office Add-ins Security
- Use HTTPS in production
- Validate all user inputs
- Implement proper authentication
- Follow Microsoft security guidelines

### API Security
- Secure Microsoft Graph API calls
- Use proper OAuth 2.0 flows
- Implement token refresh logic
- Log security events

### Data Protection
- Encrypt sensitive data
- Follow GDPR/compliance requirements
- Implement data retention policies
- Secure communication channels

## Monitoring and Maintenance

### Logging Strategy
- Use shared Logger utility
- Centralized log aggregation
- Monitor Office.js errors
- Track user interactions

### Health Checks
- Add-in availability monitoring
- API endpoint health checks
- Certificate expiration alerts
- Performance monitoring

### Updates and Versioning
- Semantic versioning for releases
- Backward compatibility considerations
- Gradual rollout strategies
- Rollback procedures

## Troubleshooting Production Issues

### Common Production Problems
1. **Certificate issues** - Monitor SSL expiration
2. **CORS errors** - Verify domain configuration
3. **API rate limits** - Implement retry logic
4. **Office version compatibility** - Test across versions

### Emergency Procedures
1. Rollback to previous version
2. Disable problematic features
3. Contact Microsoft support if needed
4. Communicate with users about issues

## Performance Optimization

### Add-in Performance
- Minimize bundle sizes
- Lazy load components
- Cache static resources
- Optimize Office.js calls

### Server Performance
- Use CDN for static assets
- Implement caching strategies
- Monitor resource usage
- Scale based on demand
