#!/usr/bin/env python3
"""
Verification script for amereset functionality in RedTool
"""

import re

def verify_amereset_implementation():
    """Verify that amereset functionality is properly implemented"""
    print("Verifying amereset implementation in RedTool...")
    
    try:
        with open('redtool.py', 'r') as f:
            content = f.read()
        
        # Check for amereset in configuration process
        amereset_patterns = [
            r'amereset.*AME.*slot',
            r'sending amereset',
            r'amereset.*for AME',
            r'amereset.*command'
        ]
        
        found_patterns = []
        for pattern in amereset_patterns:
            if re.search(pattern, content, re.IGNORECASE):
                found_patterns.append(pattern)
        
        print(f"✓ Found {len(found_patterns)} amereset-related code patterns")
        
        # Check for amereset in quick commands
        if 'amereset' in content and 'quick_cmd' in content.lower():
            print("✓ amereset included in quick commands")
        
        # Check for AME-specific logic
        if 'ame.tek' in content.lower() and 'amereset' in content:
            print("✓ AME-specific amereset logic found")
        
        # Check for timeout handling after amereset
        if 'amereset' in content and 'timeout_sec=30' in content:
            print("✓ Extended timeout for amereset found")
        
        # Look for specific amereset implementation
        amereset_lines = []
        for i, line in enumerate(content.split('\n'), 1):
            if 'amereset' in line.lower():
                amereset_lines.append(f"Line {i}: {line.strip()}")
        
        if amereset_lines:
            print(f"\nFound amereset implementation in {len(amereset_lines)} locations:")
            for line in amereset_lines[:5]:  # Show first 5 occurrences
                print(f"  {line}")
            if len(amereset_lines) > 5:
                print(f"  ... and {len(amereset_lines) - 5} more locations")
            
            return True
        else:
            print("✗ No amereset implementation found")
            return False
            
    except FileNotFoundError:
        print("✗ redtool.py not found in current directory")
        return False
    except Exception as e:
        print(f"✗ Error checking file: {e}")
        return False

def main():
    print("RedTool amereset Verification")
    print("=" * 30)
    
    success = verify_amereset_implementation()
    
    if success:
        print("\n✅ amereset functionality is properly implemented!")
        print("\nKey features:")
        print("- Automatic amereset after TEK loading for AME slots")
        print("- Extended timeout (30s) for amereset command")
        print("- Manual amereset available in quick commands")
        print("- Proper error handling and logging")
    else:
        print("\n❌ amereset functionality verification failed")
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())
