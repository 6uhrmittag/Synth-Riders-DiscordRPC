# Mock pywin32 module
class client:
    class Dispatch:
        def __init__(self, *args, **kwargs):
            pass
        
        def CreateShortcut(self, *args, **kwargs):
            class MockShortcut:
                def __init__(self):
                    self.TargetPath = ''
                
                def Save(self):
                    pass
            
            return MockShortcut()
