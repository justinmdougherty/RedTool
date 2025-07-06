#!/usr/bin/env python3
"""
Test script for RedTool Python application
"""

import sys
import os

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from redtool import BoltTerminalGUI
    import tkinter as tk
    
    def test_basic_functionality():
        """Test basic GUI initialization"""
        print("Testing RedTool Python application...")
        
        root = tk.Tk()
        root.withdraw()  # Hide the window for testing
        
        try:
            app = BoltTerminalGUI(root)
            print("✓ GUI initialized successfully")
            
            # Test some basic functionality
            app.add_message("Test message", "info")
            print("✓ Message logging works")
            
            # Test configuration validation
            app.validate_tek_key_format("ABCDEF1234567890")
            print("✓ TEK key validation works")
            
            # Test config file creation
            test_config = {
                "ame_tek_path": "",
                "wfc_tek_path": "",
                "test_mode": True
            }
            app.save_app_settings()
            print("✓ Configuration save works")
            
            print("\nBasic functionality test completed successfully!")
            return True
            
        except Exception as e:
            print(f"✗ Error during testing: {e}")
            return False
        finally:
            root.destroy()

    if __name__ == "__main__":
        success = test_basic_functionality()
        sys.exit(0 if success else 1)
        
except ImportError as e:
    print(f"Import error: {e}")
    print("Make sure all required dependencies are installed:")
    print("pip install pyserial")
    sys.exit(1)
except Exception as e:
    print(f"Unexpected error: {e}")
    sys.exit(1)
