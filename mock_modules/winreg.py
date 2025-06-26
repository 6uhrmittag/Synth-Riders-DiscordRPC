# Mock winreg module
HKEY_CURRENT_USER = 0
KEY_SET_VALUE = 0
REG_SZ = 0

def OpenKey(*args, **kwargs):
    class MockKey:
        def __enter__(self):
            return self
        
        def __exit__(self, *args):
            pass
    
    return MockKey()

def SetValueEx(*args, **kwargs):
    pass
