#!/usr/bin/env python3
"""
Local LLM Server Management
Manages local LLM models using Ollama or similar backends
"""

import os
import subprocess
import time
import requests
import psutil
from pathlib import Path
from loguru import logger
from dotenv import load_dotenv
import json

# Load environment variables
load_dotenv()

class LocalLLMManager:
    def __init__(self):
        self.ollama_host = os.getenv("OLLAMA_HOST", "127.0.0.1")
        self.ollama_port = int(os.getenv("OLLAMA_PORT", 11434))
        self.ollama_url = f"http://{self.ollama_host}:{self.ollama_port}"
        self.default_models = [
            "llama2",           # Good general purpose model
            "mistral",          # Fast and capable
            "nomic-embed-text"  # For embeddings
        ]
        
        # Configure logging
        logger.add("logs/local-llm.log", rotation="500 MB", retention="10 days")
    
    def is_ollama_running(self) -> bool:
        """Check if Ollama server is running"""
        try:
            response = requests.get(f"{self.ollama_url}/api/version", timeout=5)
            return response.status_code == 200
        except requests.exceptions.RequestException:
            return False
    
    def start_ollama(self) -> bool:
        """Start Ollama server"""
        if self.is_ollama_running():
            logger.info("Ollama is already running")
            return True
        
        try:
            logger.info("Starting Ollama server...")
            
            # Start Ollama in background
            if os.name == 'nt':  # Windows
                subprocess.Popen(
                    ["ollama", "serve"],
                    creationflags=subprocess.CREATE_NEW_CONSOLE
                )
            else:  # Unix-like
                subprocess.Popen(
                    ["ollama", "serve"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
            
            # Wait for server to start
            for _ in range(30):  # Wait up to 30 seconds
                time.sleep(1)
                if self.is_ollama_running():
                    logger.info("Ollama server started successfully")
                    return True
            
            logger.error("Failed to start Ollama server")
            return False
            
        except FileNotFoundError:
            logger.error("Ollama not found. Please install Ollama first.")
            return False
        except Exception as e:
            logger.error(f"Error starting Ollama: {e}")
            return False
    
    def stop_ollama(self) -> bool:
        """Stop Ollama server"""
        try:
            # Find and kill Ollama processes
            for proc in psutil.process_iter(['pid', 'name']):
                if 'ollama' in proc.info['name'].lower():
                    proc.kill()
                    logger.info(f"Stopped Ollama process {proc.info['pid']}")
            return True
        except Exception as e:
            logger.error(f"Error stopping Ollama: {e}")
            return False
    
    def list_models(self) -> list:
        """List available models"""
        try:
            response = requests.get(f"{self.ollama_url}/api/tags")
            if response.status_code == 200:
                data = response.json()
                return [model['name'] for model in data.get('models', [])]
            return []
        except Exception as e:
            logger.error(f"Error listing models: {e}")
            return []
    
    def pull_model(self, model_name: str) -> bool:
        """Download/pull a model"""
        try:
            logger.info(f"Pulling model: {model_name}")
            
            response = requests.post(
                f"{self.ollama_url}/api/pull",
                json={"name": model_name},
                stream=True
            )
            
            for line in response.iter_lines():
                if line:
                    data = json.loads(line)
                    if data.get('status'):
                        logger.info(f"Pull progress: {data['status']}")
                    if data.get('error'):
                        logger.error(f"Pull error: {data['error']}")
                        return False
            
            logger.info(f"Successfully pulled model: {model_name}")
            return True
            
        except Exception as e:
            logger.error(f"Error pulling model {model_name}: {e}")
            return False
    
    def setup_default_models(self):
        """Setup default models for the project"""
        logger.info("Setting up default models...")
        
        available_models = self.list_models()
        
        for model in self.default_models:
            if model not in available_models:
                logger.info(f"Model {model} not found, downloading...")
                self.pull_model(model)
            else:
                logger.info(f"Model {model} already available")
    
    def get_status(self) -> dict:
        """Get server status and model information"""
        status = {
            "running": self.is_ollama_running(),
            "url": self.ollama_url,
            "models": self.list_models() if self.is_ollama_running() else []
        }
        
        if status["running"]:
            try:
                # Get version info
                response = requests.get(f"{self.ollama_url}/api/version")
                if response.status_code == 200:
                    status["version"] = response.json()
            except:
                pass
        
        return status
    
    def test_model(self, model_name: str = "llama2") -> bool:
        """Test if a model is working"""
        try:
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": model_name,
                    "prompt": "Hello, world!",
                    "stream": False
                },
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                logger.info(f"Model {model_name} test successful: {data.get('response', '')[:50]}...")
                return True
            else:
                logger.error(f"Model {model_name} test failed: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"Error testing model {model_name}: {e}")
            return False

def main():
    """Main function for CLI usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Local LLM Server Manager")
    parser.add_argument("command", choices=["start", "stop", "status", "setup", "test"])
    parser.add_argument("--model", help="Model name for test command")
    
    args = parser.parse_args()
    
    # Create logs directory
    os.makedirs("logs", exist_ok=True)
    
    manager = LocalLLMManager()
    
    if args.command == "start":
        if manager.start_ollama():
            print("✅ Ollama server started successfully")
        else:
            print("❌ Failed to start Ollama server")
    
    elif args.command == "stop":
        if manager.stop_ollama():
            print("✅ Ollama server stopped")
        else:
            print("❌ Failed to stop Ollama server")
    
    elif args.command == "status":
        status = manager.get_status()
        print(f"Server running: {'✅' if status['running'] else '❌'}")
        print(f"URL: {status['url']}")
        print(f"Available models: {', '.join(status['models']) or 'None'}")
        if status.get('version'):
            print(f"Version: {status['version']}")
    
    elif args.command == "setup":
        if not manager.is_ollama_running():
            print("Starting Ollama server first...")
            manager.start_ollama()
        
        print("Setting up default models...")
        manager.setup_default_models()
        print("✅ Setup complete")
    
    elif args.command == "test":
        model = args.model or "llama2"
        if manager.test_model(model):
            print(f"✅ Model {model} is working")
        else:
            print(f"❌ Model {model} test failed")

if __name__ == "__main__":
    main()
