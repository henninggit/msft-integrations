#!/usr/bin/env node

const { execSync } = require('child_process');
const path = require('path');

/**
 * Development script for Microsoft Integrations Monolith
 * Provides easy access to common development tasks
 */

const projectRoot = path.dirname(__dirname);

const commands = {
    'word-dev': {
        description: 'Start Word Add-in development server',
        action: () => {
            const wordPath = path.join(projectRoot, 'integrations', 'word-addin');
            execSync('npm run dev', { cwd: wordPath, stdio: 'inherit' });
        }
    },
    'word-start': {
        description: 'Start Word Add-in for sideloading',
        action: () => {
            const wordPath = path.join(projectRoot, 'integrations', 'word-addin');
            execSync('npm start', { cwd: wordPath, stdio: 'inherit' });
        }
    },
    'llm-gateway': {
        description: 'Start LLM Gateway server',
        action: () => {
            const gatewayPath = path.join(projectRoot, 'ai-services', 'llm-gateway');
            execSync('python server.py', { cwd: gatewayPath, stdio: 'inherit' });
        }
    },
    'local-llm': {
        description: 'Start Local LLM server',
        action: () => {
            const llmPath = path.join(projectRoot, 'ai-services', 'local-llm');
            execSync('python server.py start', { cwd: llmPath, stdio: 'inherit' });
        }
    },
    'ai-setup': {
        description: 'Setup AI services (install dependencies and models)',
        action: () => {
            const setupScript = path.join(projectRoot, 'scripts', 'setup-ai.py');
            execSync(`python "${setupScript}"`, { stdio: 'inherit' });
        }
    },
    'ai-status': {
        description: 'Check AI services status',
        action: () => {
            const llmPath = path.join(projectRoot, 'ai-services', 'local-llm');
            execSync('python server.py status', { cwd: llmPath, stdio: 'inherit' });
        }
    },
    'setup': {
        description: 'Install all dependencies across the monolith',
        action: () => {
            const setupScript = path.join(projectRoot, 'scripts', 'setup-all.js');
            execSync(`node "${setupScript}"`, { stdio: 'inherit' });
        }
    },
    'help': {
        description: 'Show this help message',
        action: showHelp
    }
};

function showHelp() {
    console.log('Microsoft Integrations Monolith - Development Commands\n');
    console.log('Usage: npm run <command>\n');
    console.log('Available commands:');
    
    Object.entries(commands).forEach(([cmd, { description }]) => {
        console.log(`  ${cmd.padEnd(12)} ${description}`);
    });
    
    console.log('\nProject Structure:');
    console.log('  integrations/');
    console.log('    word-addin/        - Word Office Add-in');
    console.log('    excel-addin/       - Excel integration (planned)');
    console.log('    powerpoint-addin/  - PowerPoint integration (planned)');
    console.log('    teams-app/         - Teams application (planned)');
    console.log('    outlook-addin/     - Outlook add-in (planned)');
    console.log('  ai-services/');
    console.log('    llm-gateway/       - LiteLLM proxy server');
    console.log('    local-llm/         - Local model server (Ollama)');
    console.log('    shared/            - AI utilities and helpers');
    
    console.log('\nQuick Start:');
    console.log('  1. npm run setup          - Install all dependencies');
    console.log('  2. npm run ai-setup       - Setup AI services');
    console.log('  3. npm run llm-gateway    - Start LLM gateway');
    console.log('  4. npm run word-dev       - Start Word add-in development');
}

// Parse command line arguments
const command = process.argv[2];

if (!command || !commands[command]) {
    console.error(`Unknown command: ${command || 'none'}\n`);
    showHelp();
    process.exit(1);
}

// Execute the command
try {
    commands[command].action();
} catch (error) {
    console.error(`Error executing ${command}:`, error.message);
    process.exit(1);
}
