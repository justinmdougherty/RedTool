#!/usr/bin/env python3
"""
RedTool Launcher Script
Checks dependencies and starts the main application
"""

import sys
import subprocess
import os

def check_python_version():
    """Check if Python version is suitable"""
    if sys.version_info < (3, 7):
        print("ERROR: Python 3.7 or higher is required")
        print(f"Current version: {sys.version}")
        return False
    return True

def check_dependencies():
    """Check and install required dependencies"""
    try:
        import serial
        print("✓ pyserial is available")
        return True
    except ImportError:
        print("Installing pyserial...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyserial>=3.5"])
            print("✓ pyserial installed successfully")
            return True
        except subprocess.CalledProcessError:
            print("ERROR: Failed to install pyserial")
            print("Please run: pip install pyserial")
            return False

def main():
    """Main launcher function"""
    print("RedTool - BOLT Terminal & Configurator")
    print("=" * 40)
    
    # Check Python version
    if not check_python_version():
        return 1
        
    print(f"✓ Python {sys.version.split()[0]} found")
    
    # Check dependencies
    if not check_dependencies():
        return 1
    
    # Change to script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Import and start the application
    try:
        print("Starting RedTool...")
        print()
        
        from redtool import main as redtool_main
        redtool_main()
        
    except ImportError as e:
        print(f"ERROR: Failed to import redtool: {e}")
        print("Make sure redtool.py is in the same directory")
        return 1
    except KeyboardInterrupt:
        print("\nApplication interrupted by user")
        return 0
    except Exception as e:
        print(f"ERROR: Application failed to start: {e}")
        return 1
    
    return 0

if __name__ == "__main__":
    exit_code = main()
    if exit_code != 0:
        input("Press Enter to continue...")
    sys.exit(exit_code)
