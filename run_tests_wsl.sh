#!/bin/bash

# Script to run tests in Ubuntu WSL
# This script helps run the tests in a Linux environment by mocking Windows-specific dependencies

echo "Setting up environment for running tests in WSL..."

# Create a virtual environment (optional but recommended)
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate the virtual environment
source venv/bin/activate

# Install dependencies (excluding Windows-specific ones)
echo "Installing dependencies..."
pip install pypresence psutil rich websocket-client requests
pip install pytest flake8 black mypy

# Create mock modules directory
echo "Creating mock modules for Windows-specific dependencies..."
mkdir -p mock_modules

# Create mock win32com package
mkdir -p mock_modules/win32com
cat > mock_modules/win32com/__init__.py << 'EOF'
# Mock win32com package
EOF

cat > mock_modules/win32com/client.py << 'EOF'
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
EOF

# Create mock winreg module
cat > mock_modules/winreg.py << 'EOF'
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
EOF

# Set PYTHONPATH to include mock modules
export PYTHONPATH=$PYTHONPATH:$(pwd)/mock_modules

# Run the tests
echo "Running tests..."
pytest

# Run linting and type checking
echo "Running linting with flake8..."
flake8 src/ --count --select=E9,F63,F7,F82 --show-source --statistics
flake8 src/ --count --exit-zero --max-complexity=10 --max-line-length=127 --statistics

echo "Checking formatting with black..."
black --check src/

echo "Running type checking with mypy..."
mypy --ignore-missing-imports src/

echo "Done!"
