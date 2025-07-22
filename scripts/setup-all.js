#!/usr/bin/env node

const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

/**
 * Setup script for Microsoft Integrations Monolith
 * Installs dependencies for all integration projects
 */

const projectRoot = path.dirname(__dirname);
const integrationsPath = path.join(projectRoot, 'integrations');

console.log('🚀 Setting up Microsoft Integrations Monolith...\n');

// Install root dependencies
console.log('📦 Installing root dependencies...');
try {
    execSync('npm install', { 
        cwd: projectRoot, 
        stdio: 'inherit' 
    });
    console.log('✅ Root dependencies installed\n');
} catch (error) {
    console.error('❌ Failed to install root dependencies:', error.message);
    process.exit(1);
}

// Find and install dependencies for all integration projects
if (fs.existsSync(integrationsPath)) {
    const integrations = fs.readdirSync(integrationsPath, { withFileTypes: true })
        .filter(dirent => dirent.isDirectory())
        .map(dirent => dirent.name);

    for (const integration of integrations) {
        const integrationPath = path.join(integrationsPath, integration);
        const packageJsonPath = path.join(integrationPath, 'package.json');
        
        if (fs.existsSync(packageJsonPath)) {
            console.log(`📦 Installing dependencies for ${integration}...`);
            try {
                execSync('npm install', { 
                    cwd: integrationPath, 
                    stdio: 'inherit' 
                });
                console.log(`✅ ${integration} dependencies installed\n`);
            } catch (error) {
                console.error(`❌ Failed to install ${integration} dependencies:`, error.message);
                process.exit(1);
            }
        }
    }
}

console.log('🎉 All dependencies installed successfully!');
console.log('\nNext steps:');
console.log('  • For Word Add-in development: npm run word-dev');
console.log('  • To start Word Add-in: npm run word-start');
console.log('  • For more options: npm run help');
