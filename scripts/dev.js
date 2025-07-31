#!/usr/bin/env node

const { execSync } = require('child_process');
const path = require('path');

/**
 * Development script for Microsoft Integrations Monolith
 * Provides easy access to common development tasks
 */

const projectRoot = path.dirname(__dirname);

const commands = {
    'word-webpack': {
        description: 'Start Word webpack',
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
            const venvPython = path.join(projectRoot, 'ai-services', 'venv', 'Scripts', 'python.exe');
            execSync(`"${venvPython}" server.py`, { cwd: gatewayPath, stdio: 'inherit' });
        }
    },
    'local-llm': {
        description: 'Start Local LLM server',
        action: () => {
            const llmPath = path.join(projectRoot, 'ai-services', 'local-llm');
            const venvPython = path.join(projectRoot, 'ai-services', 'venv', 'Scripts', 'python.exe');
            execSync(`"${venvPython}" server.py start`, { cwd: llmPath, stdio: 'inherit' });
        }
    },
    'ai-setup': {
        description: 'Setup AI services (install dependencies and models)',
        action: () => {
            const setupScript = path.join(projectRoot, 'scripts', 'setup-ai.py');
            const venvPython = path.join(projectRoot, 'ai-services', 'venv', 'Scripts', 'python.exe');
            execSync(`"${venvPython}" "${setupScript}"`, { stdio: 'inherit' });
        }
    },
    'ai-status': {
        description: 'Check AI services status',
        action: () => {
            const llmPath = path.join(projectRoot, 'ai-services', 'local-llm');
            const venvPython = path.join(projectRoot, 'ai-services', 'venv', 'Scripts', 'python.exe');
            execSync(`"${venvPython}" server.py status`, { cwd: llmPath, stdio: 'inherit' });
        }
    },
    'setup': {
        description: 'Install all dependencies across the monolith',
        action: () => {
            const setupScript = path.join(projectRoot, 'scripts', 'setup-all.js');
            execSync(`node "${setupScript}"`, { stdio: 'inherit' });
        }
    },
    'stop': {
        description: 'Stop all running project services',
        action: stopAllServices
    },
    'help': {
        description: 'Show this help message',
        action: showHelp
    }
};

function stopAllServices() {
    console.log('🛑 Stopping all Microsoft Integrations services...\n');
    
    try {
        // Kill processes by port (common development ports)
        const ports = [3000, 8000, 11434]; // webpack dev server, LLM gateway, Ollama
        
        ports.forEach(port => {
            try {
                console.log(`🔍 Checking port ${port}...`);
                // Use PowerShell to find and kill processes on specific ports
                execSync(`powershell -Command "Get-NetTCPConnection -LocalPort ${port} -ErrorAction SilentlyContinue | ForEach-Object { Stop-Process -Id (Get-Process -Id $_.OwningProcess).Id -Force }"`, { stdio: 'inherit' });
                console.log(`✅ Stopped services on port ${port}`);
            } catch (error) {
                console.log(`ℹ️  No services running on port ${port}`);
            }
        });
        
        // Kill processes by name (common Node.js and Python processes)
        const processNames = ['node.exe', 'python.exe', 'uvicorn.exe'];
        
        processNames.forEach(processName => {
            try {
                console.log(`🔍 Stopping ${processName} processes...`);
                execSync(`taskkill /F /IM "${processName}" /T 2>nul`, { stdio: 'pipe' });
                console.log(`✅ Stopped ${processName} processes`);
            } catch (error) {
                console.log(`ℹ️  No ${processName} processes found`);
            }
        });
        
        console.log('\n🎉 All services stopped successfully!');
        console.log('💡 You can restart services with:');
        console.log('   npm run llm-gateway    - Start LLM Gateway');
        console.log('   npm run word-webpack   - Start Word development');
        console.log('   npm run local-llm      - Start Local LLM (requires Ollama)');
        
    } catch (error) {
        console.error('❌ Error stopping services:', error.message);
        console.log('\n💡 Manual cleanup:');
        console.log('   - Press Ctrl+C in any running terminal windows');
        console.log('   - Check Task Manager for remaining processes');
    }
}

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
    console.log('  4. npm run word-webpack   - Start Word add-in development');
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
