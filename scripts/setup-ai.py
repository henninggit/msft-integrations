#!/usr/bin/env python3
"""
AI Services Setup Script
Automates the setup process for AI components
"""

import os
import sys
import subprocess
import platform
import requests
import time
from pathlib import Path

def print_header(title):
    print(f"\n{'='*60}")
    print(f" {title}")
    print(f"{'='*60}")

def print_step(step):
    print(f"\n🔧 {step}")

def print_success(message):
    print(f"✅ {message}")

def print_error(message):
    print(f"❌ {message}")

def print_warning(message):
    print(f"⚠️  {message}")

def run_command(command, cwd=None, check=True):
    """Run a command and return success status"""
    try:
        result = subprocess.run(
            command,
            shell=True,
            cwd=cwd,
            capture_output=True,
            text=True,
            check=check
        )
        return True, result.stdout, result.stderr
    except subprocess.CalledProcessError as e:
        return False, e.stdout, e.stderr

def check_python():
    """Check if Python is available and version >= 3.8"""
    print_step("Checking Python installation...")
    
    try:
        version = sys.version_info
        if version.major == 3 and version.minor >= 8:
            print_success(f"Python {version.major}.{version.minor}.{version.micro} found")
            return True
        else:
            print_error(f"Python 3.8+ required, found {version.major}.{version.minor}")
            return False
    except:
        print_error("Python not found")
        return False

def check_pip():
    """Check if pip is available"""
    print_step("Checking pip installation...")
    
    success, stdout, stderr = run_command("pip --version")
    if success:
        print_success("pip is available")
        return True
    else:
        print_error("pip not found")
        return False

def install_python_dependencies():
    """Install Python dependencies for AI services"""
    print_step("Installing Python dependencies...")
    
    # Get the directory of this script
    script_dir = Path(__file__).parent
    ai_services_dir = script_dir.parent / "ai-services"
    
    # Install LLM Gateway dependencies
    gateway_dir = ai_services_dir / "llm-gateway"
    requirements_file = gateway_dir / "requirements.txt"
    
    if requirements_file.exists():
        print("Installing LLM Gateway dependencies...")
        success, stdout, stderr = run_command(
            f"pip install -r {requirements_file}",
            cwd=str(gateway_dir)
        )
        if success:
            print_success("LLM Gateway dependencies installed")
        else:
            print_error(f"Failed to install LLM Gateway dependencies: {stderr}")
            return False
    
    # Install Local LLM dependencies
    llm_dir = ai_services_dir / "local-llm"
    requirements_file = llm_dir / "requirements.txt"
    
    if requirements_file.exists():
        print("Installing Local LLM dependencies...")
        success, stdout, stderr = run_command(
            f"pip install -r {requirements_file}",
            cwd=str(llm_dir)
        )
        if success:
            print_success("Local LLM dependencies installed")
        else:
            print_error(f"Failed to install Local LLM dependencies: {stderr}")
            return False
    
    return True

def check_ollama():
    """Check if Ollama is installed"""
    print_step("Checking Ollama installation...")
    
    success, stdout, stderr = run_command("ollama --version", check=False)
    if success:
        print_success("Ollama is installed")
        return True
    else:
        print_warning("Ollama not found")
        return False

def install_ollama():
    """Install Ollama"""
    print_step("Installing Ollama...")
    
    system = platform.system().lower()
    
    if system == "windows":
        print("Please download and install Ollama from: https://ollama.ai/download/windows")
        print("After installation, restart this script.")
        return False
    elif system == "darwin":  # macOS
        success, stdout, stderr = run_command("brew install ollama", check=False)
        if success:
            print_success("Ollama installed via Homebrew")
            return True
        else:
            print("Please download and install Ollama from: https://ollama.ai/download/mac")
            return False
    else:  # Linux
        success, stdout, stderr = run_command(
            "curl -fsSL https://ollama.ai/install.sh | sh",
            check=False
        )
        if success:
            print_success("Ollama installed")
            return True
        else:
            print_error("Failed to install Ollama")
            return False

def setup_environment():
    """Setup environment file"""
    print_step("Setting up environment configuration...")
    
    script_dir = Path(__file__).parent
    ai_services_dir = script_dir.parent / "ai-services"
    
    env_example = ai_services_dir / ".env.example"
    env_file = ai_services_dir / ".env"
    
    if env_example.exists() and not env_file.exists():
        # Copy example to .env
        with open(env_example, 'r') as f:
            content = f.read()
        
        with open(env_file, 'w') as f:
            f.write(content)
        
        print_success("Environment file created")
        print("You can customize settings in ai-services/.env")
    else:
        print_success("Environment file already exists")
    
    return True

def start_ollama():
    """Start Ollama server"""
    print_step("Starting Ollama server...")
    
    # Check if already running
    try:
        response = requests.get("http://localhost:11434/api/version", timeout=5)
        if response.status_code == 200:
            print_success("Ollama is already running")
            return True
    except:
        pass
    
    # Start Ollama
    if platform.system().lower() == "windows":
        subprocess.Popen(["ollama", "serve"], creationflags=subprocess.CREATE_NEW_CONSOLE)
    else:
        subprocess.Popen(["ollama", "serve"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    # Wait for server to start
    for i in range(30):
        time.sleep(1)
        try:
            response = requests.get("http://localhost:11434/api/version", timeout=5)
            if response.status_code == 200:
                print_success("Ollama server started")
                return True
        except:
            continue
    
    print_error("Failed to start Ollama server")
    return False

def pull_default_models():
    """Pull default models"""
    print_step("Pulling default models...")
    
    models = ["llama2", "mistral"]
    
    for model in models:
        print(f"Pulling {model}...")
        success, stdout, stderr = run_command(f"ollama pull {model}")
        if success:
            print_success(f"{model} pulled successfully")
        else:
            print_warning(f"Failed to pull {model}: {stderr}")
    
    return True

def test_setup():
    """Test the complete setup"""
    print_step("Testing setup...")
    
    # Test Ollama
    try:
        response = requests.get("http://localhost:11434/api/version", timeout=5)
        if response.status_code == 200:
            print_success("Ollama server is responding")
        else:
            print_warning("Ollama server not responding properly")
    except:
        print_warning("Cannot connect to Ollama server")
    
    # Test a simple generation
    try:
        test_data = {
            "model": "llama2",
            "prompt": "Hello",
            "stream": False
        }
        response = requests.post(
            "http://localhost:11434/api/generate",
            json=test_data,
            timeout=30
        )
        if response.status_code == 200:
            print_success("Model generation test passed")
        else:
            print_warning("Model generation test failed")
    except:
        print_warning("Model generation test failed")

def main():
    print_header("AI Services Setup")
    print("This script will set up the AI services for the Microsoft Integrations project.")
    
    # Check prerequisites
    if not check_python():
        print_error("Python 3.8+ is required")
        sys.exit(1)
    
    if not check_pip():
        print_error("pip is required")
        sys.exit(1)
    
    # Install Python dependencies
    if not install_python_dependencies():
        print_error("Failed to install Python dependencies")
        sys.exit(1)
    
    # Check/install Ollama
    if not check_ollama():
        print("Ollama is required for local LLM functionality.")
        install = input("Would you like to install Ollama? (y/N): ").strip().lower()
        if install == 'y':
            if not install_ollama():
                print_error("Failed to install Ollama")
                sys.exit(1)
        else:
            print_warning("Skipping Ollama installation")
    
    # Setup environment
    setup_environment()
    
    # Start Ollama and pull models
    if check_ollama():
        if start_ollama():
            pull_default_models()
        
        # Test the setup
        test_setup()
    
    print_header("Setup Complete!")
    print("Next steps:")
    print("1. Start the LLM Gateway: npm run llm-gateway")
    print("2. Start Word add-in development: npm run word-dev")
    print("3. Test AI features in your Office applications")
    print("\nFor more information, see ai-services/README.md")

if __name__ == "__main__":
    main()
