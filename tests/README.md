# Synth Riders DiscordRPC Tests

This directory contains automated tests for the Synth Riders DiscordRPC project.

## Overview

The test suite focuses on testing the core logic of the application, including:

- Configuration handling (`config.json` load/save/migrate)
- Setup logic (e.g. first-run directory creation, defaults)
- Utility logic (e.g. formatting, parsing)
- Album art upload logic (base64 decode → POST → URL extract)
- Discord RPC integration (mock `pypresence.Presence.update`)
- WebSocket event handling (mock events like `SongStart`, `SongEnd`)

## Running Tests

### On Windows

To run the tests on Windows, you'll need to install the test dependencies:

```bash
pip install pytest flake8 black mypy
```

Then, run the tests with:

```bash
pytest
```

### On Ubuntu WSL

To run the tests on Ubuntu WSL, use the provided script:

```bash
# Make the script executable
chmod +x run_tests_wsl.sh

# Run the script
./run_tests_wsl.sh
```

This script will:
1. Create a virtual environment (if it doesn't exist)
2. Install the necessary dependencies (excluding Windows-specific ones)
3. Create mock modules for Windows-specific imports
4. Set up the environment
5. Run the tests, linting, and type checking

The script ensures that the same tests that run in the CI workflow can run in your Ubuntu WSL environment.

## Test Files

- `test_config.py`: Tests for configuration handling
- `test_setup.py`: Tests for setup logic
- `test_utilities.py`: Tests for utility functions and album art upload
- `test_presence.py`: Tests for Discord RPC integration and WebSocket event handling
- `conftest.py`: Pytest configuration and fixtures

## CI Integration

This project uses GitHub Actions for continuous integration. The workflow is defined in `.github/workflows/ci.yml` and runs the tests on Ubuntu Linux.

The CI workflow:
1. Runs on pushes to the master branch and pull requests
2. Uses Python 3.11
3. Installs dependencies
4. Mocks Windows-specific dependencies
5. Runs pytest, flake8, black, and mypy

## Notes for Contributors

When adding new features, please also add corresponding tests. The tests should:

- Be platform-independent (avoid Windows-specific features)
- Use mocks for external dependencies (Discord, WebSocket connections)
- Cover both success and error cases
- Be clear and concise

Remember that the application is primarily designed for Windows, but the core logic should be testable on any platform.
