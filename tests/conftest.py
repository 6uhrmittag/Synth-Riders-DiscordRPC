"""
Configuration for pytest.

This file contains fixtures and configuration for pytest.
"""
import os
import sys
import pytest
from unittest.mock import patch

# Add mock modules for Windows-specific imports when running on non-Windows platforms
@pytest.fixture(autouse=True)
def mock_windows_modules():
    """Mock Windows-specific modules when running on non-Windows platforms."""
    if sys.platform != 'win32':
        # Create mock modules for Windows-specific imports
        modules_to_mock = {
            'win32com': True,
            'win32com.client': True,
            'winreg': True,
        }
        
        for module_name in modules_to_mock:
            if module_name not in sys.modules:
                sys.modules[module_name] = type(module_name, (), {})
                
        # Mock specific classes and functions
        if 'win32com.client' in sys.modules:
            class MockDispatch:
                def __init__(self, *args, **kwargs):
                    pass
                
                def CreateShortcut(self, *args, **kwargs):
                    class MockShortcut:
                        def __init__(self):
                            self.TargetPath = ''
                        
                        def Save(self):
                            pass
                    
                    return MockShortcut()
            
            sys.modules['win32com.client'].Dispatch = MockDispatch
        
        if 'winreg' in sys.modules:
            sys.modules['winreg'].HKEY_CURRENT_USER = 0
            sys.modules['winreg'].KEY_SET_VALUE = 0
            sys.modules['winreg'].REG_SZ = 0
            
            def mock_open_key(*args, **kwargs):
                class MockKey:
                    def __enter__(self):
                        return self
                    
                    def __exit__(self, *args):
                        pass
                
                return MockKey()
            
            sys.modules['winreg'].OpenKey = mock_open_key
            sys.modules['winreg'].SetValueEx = lambda *args, **kwargs: None
        
        # Patch psutil.Process for Windows-specific process handling
        patcher = patch('psutil.Process')
        mock_process = patcher.start()
        mock_process.return_value.name.return_value = 'SynthRiders.exe'
        
        yield
        
        # Stop the patcher
        patcher.stop()
    else:
        yield