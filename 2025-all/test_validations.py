"""Quick validation tests for build-tournament-folders.py"""
import os
import json
import shutil
from pathlib import Path

def test_missing_config():
    """Test that script fails gracefully with missing config"""
    print("Testing missing config file...")
    # Your test code here
    
def test_invalid_json():
    """Test that script handles invalid JSON"""
    print("Testing invalid JSON...")
    # Your test code here

if __name__ == "__main__":
    print("Running validation tests...\n")
    test_missing_config()
    test_invalid_json()
    print("\nTests complete!")
