# Testing Guide

This directory contains all tests for the Synth Riders DiscordRPC application.

## Setup

### Install Test Dependencies

Before running tests, install the test dependencies:

```bash
pip install -r tests/requirements-test.txt
```

Or if you want both main and test dependencies:

```bash
pip install -r requirements.txt
pip install -r tests/requirements-test.txt
```

## Running Tests

### Run All Tests

```bash
pytest tests/
```

### Run with Verbose Output

```bash
pytest -v tests/
```

### Run Specific Test File

```bash
pytest tests/test_logger.py
pytest tests/test_presence_helpers.py
pytest tests/test_websocket_integration.py
```

### Run Specific Test

```bash
pytest tests/test_logger.py::TestLogger::test_logger_info
```

### Run with Coverage

```bash
pytest --cov=src tests/
```

## Test Structure

- **`test_logger.py`** - Tests for the logging functionality
  - Tests log file creation
  - Tests info, error, warning methods
  - Tests log clearing
  
- **`test_presence_helpers.py`** - Tests for Presence helper methods
  - Tests `format_time()` method for time formatting
  - Tests `synth_riders_process_exists()` for process detection
  - Tests websocket event handling
  
- **`test_websocket_integration.py`** - Integration tests with mock websocket
  - Tests mock websocket server
  - Tests event creation functions
  - Tests full integration with Presence class

- **`mock_websocket_server.py`** - Mock server simulating SynthRiders-Websockets-Mod
  - Can be used standalone for manual testing
  - Implements all events from the official mod

## Manual Testing with Mock Server

You can run the mock websocket server standalone to test the application without the actual game:

### Start the Mock Server

```bash
python tests/mock_websocket_server.py
```

This will start an interactive mock server on `ws://localhost:9000` (the default port).

### Interactive Commands

Once the server is running and a client connects, you can use these commands:

- `1` - Simulate a full song playthrough (SongStart → PlayTime → NoteHit → SongEnd)
- `2` - Simulate returning to menu mid-song
- `3` - Send a SongStart event
- `4` - Send a PlayTime event
- `5` - Send a NoteHit event
- `6` - Send a SongEnd event
- `q` - Quit

### Testing with the Real Application

1. Start the mock server: `python tests/mock_websocket_server.py`
2. Start the RPC application (the built executable or run `python src/bin/rpc.py` after setup)
3. Use the interactive commands to send events
4. Watch Discord to see the Rich Presence update

## Troubleshooting

### ImportError: No module named 'src'

Make sure you're running pytest from the project root directory:

```bash
cd /path/to/Synth-Riders-DiscordRPC
pytest tests/
```

### Tests fail with "Address already in use"

The mock websocket server port (9001 for tests, 9000 for manual testing) is already in use. Close any running instances of the mock server or the actual SynthRiders-Websockets-Mod.

### Mock server doesn't receive connections

Make sure:
1. The mock server is running before starting the client
2. The correct port is configured (9000 for manual testing)
3. No firewall is blocking localhost connections

### Tests pass but application doesn't work

The tests use mocked dependencies (Discord, file upload, etc.). To test the full application:
1. Make sure Discord is running
2. Use the mock websocket server for game events
3. Check the log file at `%localappdata%\Synth Riders DiscordRPC\logs\log.txt`

## Writing New Tests

When adding new features, please add corresponding tests:

1. Add unit tests for individual functions in `test_presence_helpers.py` or create a new test file
2. If adding new websocket events, add tests in `test_websocket_integration.py`
3. Mock external dependencies (Discord API, HTTP requests, file system)
4. Use fixtures for common setup code

### Example Test

```python
def test_new_feature(mock_presence):
    """Test description"""
    # Arrange
    mock_presence.some_value = 10
    
    # Act
    result = mock_presence.some_method()
    
    # Assert
    assert result == expected_value
```

## Continuous Integration

Tests should pass before merging any changes. Run the full test suite:

```bash
pytest -v tests/
```

All tests should pass with no errors.

