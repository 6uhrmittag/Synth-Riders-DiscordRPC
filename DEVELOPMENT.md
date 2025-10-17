# Development Guide

This guide covers local development, testing, and building the Synth Riders DiscordRPC application.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Local Setup](#local-setup)
- [Development Workflow](#development-workflow)
- [Testing](#testing)
- [Building](#building)
- [Project Structure](#project-structure)

## Prerequisites

- Python 3.10 or higher
- Git
- Discord Desktop Application (for testing Rich Presence)
- (Optional) Synth Riders VR game with [SynthRiders-Websockets-Mod](https://github.com/bookdude13/SynthRiders-Websockets-Mod)

## Local Setup

### 1. Clone the Repository

```bash
git clone https://github.com/6uhrmittag/Synth-Riders-DiscordRPC.git
cd Synth-Riders-DiscordRPC
```

### 2. Create Virtual Environment

```bash
python -m venv .venv
```

### 3. Activate Virtual Environment

**Windows (PowerShell):**
```powershell
.\.venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
.venv\Scripts\activate.bat
```

**Linux/Mac:**
```bash
source .venv/bin/activate
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
pip install -r tests/requirements-test.txt
```

## Development Workflow

### Running the Application Locally

**Important:** The application expects a specific directory structure and config file when running as an executable. For development, you can:

1. **Use the mock websocket server** (recommended for development):
   ```bash
   # Terminal 1: Start mock server
   python tests/mock_websocket_server.py
   
   # Terminal 2: Run tests to verify behavior
   pytest tests/
   ```

2. **Run with actual game:**
   - Install the SynthRiders-Websockets-Mod
   - Build and install the application (see [Building](#building))
   - Run the executable from the install location

### Testing Without Discord

The tests automatically mock the Discord/PyPresence connection, so you don't need Discord running to test the core functionality:

```bash
pytest tests/
```

### Testing Without the Game

Use the mock websocket server to simulate game events without running Synth Riders:

1. Start the mock server:
   ```bash
   python tests/mock_websocket_server.py
   ```

2. In another terminal, run the application or tests
3. Use the interactive commands to send events

The mock server simulates all events from the SynthRiders-Websockets-Mod:
- `SongStart` - Start of a song with metadata
- `PlayTime` - Song progress updates
- `NoteHit` - Score and combo updates
- `SongEnd` - End of song with results
- `ReturnToMenu` - User exits to menu
- `SceneChange` - Scene transitions

### Code Style

The project uses Python best practices:

- **Docstrings:** All functions and classes should have docstrings
- **Type hints:** Add type hints where applicable
- **Formatting:** Keep code clean and readable

Run linting:
```bash
flake8 src/ tests/
```

Format code:
```bash
black src/ tests/
```

## Testing

### Running Tests

See [tests/README.md](tests/README.md) for comprehensive testing instructions.

Quick reference:

```bash
# All tests
pytest tests/

# Verbose output
pytest -v tests/

# Specific test file
pytest tests/test_logger.py

# With coverage
pytest --cov=src tests/
```

### Manual Integration Testing

1. **Start Mock Server:**
   ```bash
   python tests/mock_websocket_server.py
   ```

2. **Build and Install Application:**
   ```bash
   build.bat
   ```
   Then run the setup executable from `dist/`

3. **Test the Application:**
   - Start the installed RPC application
   - Discord should show you as "Playing Synth Riders"
   - In the mock server terminal, press `1` to simulate a song
   - Check Discord to see the Rich Presence update with song info

4. **Check Logs:**
   Logs are written to: `%localappdata%\Synth Riders DiscordRPC\logs\log.txt`

### Testing Album Art Upload

The application uploads album art to a temporary file hosting service. For development:

- The mock server provides a test base64 PNG image
- The upload is mocked in tests
- For real testing, you need an active internet connection
- Default service: https://uguu.se/upload

## Building

### Build Executables

The project uses PyInstaller to create standalone executables:

```bash
build.bat
```

This creates three executables in the `dist/` folder:
- `Synth Riders DiscordRPC.exe` - Main application
- `Uninstall Synth Riders DiscordRPC.exe` - Uninstaller
- `Synth Riders DiscordRPC Setup.exe` - Installer/setup

### Build Process Details

The `build.bat` script runs PyInstaller with three `.spec` files:
1. `synth_riders_discordrpc.spec` - Main RPC application
2. `synth_riders_discordrpc_uninstall.spec` - Uninstaller
3. `synth_riders_discordrpc_setup.spec` - Setup/installer

### Testing Built Executables

1. Build the executables: `build.bat`
2. Run the setup: `dist\Synth Riders DiscordRPC Setup.exe`
3. Follow the installation process
4. The application will be installed to: `%localappdata%\Synth Riders DiscordRPC`
5. Test with the mock websocket server

### Version Updates

Before building a new release:

1. Update version in `config.py`:
   ```python
   VERSION = "X.Y.Z"
   ```

2. Update README.md if needed
3. Build and test
4. Commit changes
5. Tag the release:
   ```bash
   git tag -a vX.Y.Z -m "Release X.Y.Z"
   git push origin vX.Y.Z
   ```

## Project Structure

```
Synth-Riders-DiscordRPC/
├── src/
│   ├── bin/
│   │   ├── rpc.py           # Main RPC application entry point
│   │   ├── setup.py         # Installation/setup logic
│   │   └── uninstall.py     # Uninstallation logic
│   └── utilities/
│       ├── cli/             # Command-line interface utilities
│       │   ├── errors.py
│       │   ├── input.py
│       │   └── output.py
│       └── rpc/             # Rich Presence core functionality
│           ├── assets.py    # Discord asset configuration
│           ├── logger.py    # Logging functionality
│           └── presence.py  # Main presence logic and websocket handling
├── tests/
│   ├── mock_websocket_server.py  # Mock SynthRiders websocket mod
│   ├── test_logger.py            # Logger tests
│   ├── test_presence_helpers.py  # Presence helper method tests
│   ├── test_websocket_integration.py  # Integration tests
│   ├── requirements-test.txt     # Test dependencies
│   └── README.md                 # Testing guide
├── assets/                  # Images and icons
├── screenshots/             # Screenshots for README
├── config.py               # Application configuration
├── index.py                # Entry point (imports setup)
├── requirements.txt        # Python dependencies
├── build.bat              # Build script
├── *.spec                 # PyInstaller specifications
├── README.md              # User documentation
├── DEVELOPMENT.md         # This file
└── LICENSE.md             # License
```

## Key Files

### config.py
Central configuration with version, process names, websocket settings, and Discord application ID.

### src/utilities/rpc/presence.py
Core logic:
- Connects to Discord via PyPresence
- Connects to SynthRiders websocket
- Handles game events and updates Discord status
- Manages album art upload

### src/bin/setup.py
Installation logic:
- Copies executables to install location
- Creates config file
- Sets up Windows shortcuts and startup entries

## Common Development Tasks

### Adding a New Websocket Event

1. Check [SynthRiders-Websockets-Mod docs](https://github.com/bookdude13/SynthRiders-Websockets-Mod) for event structure
2. Add event handler in `src/utilities/rpc/presence.py` in `handle_websocket_event()`
3. Add event creation function in `tests/mock_websocket_server.py`
4. Add tests in `tests/test_websocket_integration.py`
5. Test with mock server

### Changing Discord Presence Format

1. Modify `update_presence()` in `src/utilities/rpc/presence.py`
2. Update tests in `tests/test_websocket_integration.py`
3. Run tests to verify
4. Build and test manually with Discord

### Modifying Configuration Options

1. Update `Config` class in `config.py`
2. Update setup process in `src/bin/setup.py` if needed
3. Update `CONFIG.md` with new option documentation
4. Update tests with new config structure

## Troubleshooting

### Application doesn't connect to Discord
- Make sure Discord Desktop is running (not web version)
- Check logs at `%localappdata%\Synth Riders DiscordRPC\logs\log.txt`
- Verify `discord_application_id` in config

### Websocket connection fails
- Check SynthRiders-Websockets-Mod is installed
- Verify port 9000 is not blocked by firewall
- Check mod's MelonPreferences.cfg for correct port
- Use mock server to test application logic

### Build fails
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Check PyInstaller version compatibility
- Look for errors in PyInstaller output

### Tests fail
- Make sure you're in the project root: `cd /path/to/Synth-Riders-DiscordRPC`
- Install test dependencies: `pip install -r tests/requirements-test.txt`
- Check for port conflicts (9001 used in tests)

## Contributing

1. Create a feature branch: `git checkout -b feature-name`
2. Make your changes
3. Add tests for new functionality
4. Run the test suite: `pytest tests/`
5. Build and test manually
6. Commit with clear messages
7. Push and create a pull request

## Resources

- [SynthRiders-Websockets-Mod](https://github.com/bookdude13/SynthRiders-Websockets-Mod) - Official websocket mod documentation
- [PyPresence](https://github.com/qwertyquerty/pypresence) - Discord Rich Presence library
- [Discord Developer Portal](https://discord.com/developers/applications) - Manage Discord applications
- [PyInstaller](https://pyinstaller.org/) - Building executables

