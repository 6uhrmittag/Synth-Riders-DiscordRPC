# Mock win32com.client module
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
