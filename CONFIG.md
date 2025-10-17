# Configuration Guide

This document explains all configuration options for the Synth Riders DiscordRPC application.

## Configuration File Location

The configuration file is automatically created during installation at:

```
%localappdata%\Synth Riders DiscordRPC\config\config.json
```

On Windows, this typically expands to:
```
C:\Users\YourUsername\AppData\Local\Synth Riders DiscordRPC\config\config.json
```

## Configuration File Structure

The `config.json` file is created during setup and contains the following fields:

```json
{
    "version": "1.1.1",
    "synthriders_install_location": "C:\\Program Files (x86)\\Steam\\steamapps\\common\\SynthRiders",
    "rich_presence_install_location": "C:\\Users\\YourUsername\\AppData\\Local\\Synth Riders DiscordRPC",
    "startup_preference": true,
    "keep_running_preference": true,
    "shortcut_preference": true,
    "promote_preference": false,
    "discord_application_id": "1124356298578870333",
    "discord_application_logo_large": "https://raw.githubusercontent.com/6uhrmittag/Synth-Riders-DiscordRPC/refs/heads/master/assets/game_synthriders_logo_alt.png",
    "discord_application_logo_small": "https://raw.githubusercontent.com/6uhrmittag/Synth-Riders-DiscordRPC/refs/heads/master/assets/game_synthriders_logo.png",
    "synthriders_websocket_host": "localhost",
    "synthriders_websocket_port": "9000",
    "image_upload_url": "https://uguu.se/upload"
}
```

## Configuration Options

### version
- **Type:** String
- **Default:** Current application version
- **Description:** The version of the application that created this config file. Used for compatibility checking.
- **Editable:** No (automatically managed)

### synthriders_install_location
- **Type:** String (file path)
- **Default:** `C:\Program Files (x86)\Steam\steamapps\common\SynthRiders`
- **Description:** The installation directory of Synth Riders VR.
- **Editable:** Yes, but must point to a valid Synth Riders installation

### rich_presence_install_location
- **Type:** String (file path)
- **Default:** `%localappdata%\Synth Riders DiscordRPC`
- **Description:** Where the DiscordRPC application is installed.
- **Editable:** No (changing this will break the installation)

### startup_preference
- **Type:** Boolean
- **Default:** Chosen during setup
- **Description:** If `true`, the application will automatically start when you log into Windows.
- **Editable:** Yes
- **Note:** Requires re-running setup to take effect, or manually modify Windows startup settings

### keep_running_preference
- **Type:** Boolean
- **Default:** Chosen during setup
- **Description:** If `true`, the application will keep running in the background even when Synth Riders is closed, and automatically reconnect when the game starts again.
- **Editable:** Yes
- **Recommended:** `true` for seamless experience

### shortcut_preference
- **Type:** Boolean
- **Default:** Chosen during setup
- **Description:** If `true`, a desktop shortcut was created during installation.
- **Editable:** No (this is just a record of installation choices)

### promote_preference
- **Type:** Boolean
- **Default:** Chosen during setup
- **Description:** If `true`, adds a "Want this status too?" button to your Discord presence linking to the GitHub repository.
- **Editable:** Yes
- **Effect:** Changes take effect on next presence update (within 15 seconds)

### discord_application_id
- **Type:** String
- **Default:** `"1124356298578870333"` (Official Synth Riders SteamVR App)
- **Alternative:** `"1342397301687189544"` (Custom App)
- **Description:** The Discord Application ID used for Rich Presence.
- **Options:**
  - **Official App (default):** Uses Synth Riders' official Discord app. Requires URLs for images (cannot upload custom images to Discord Developer Portal).
  - **Custom App:** Uses a custom Discord application. Allows custom images but requires uploading them to Discord Developer Portal first.
- **Editable:** Yes, but requires corresponding image configuration changes

### discord_application_logo_large
- **Type:** String (URL or asset name)
- **Default:** `"https://raw.githubusercontent.com/6uhrmittag/Synth-Riders-DiscordRPC/refs/heads/master/assets/game_synthriders_logo_alt.png"`
- **Description:** The large logo shown in Discord Rich Presence (background image).
- **Format:**
  - For official app: Must be a URL
  - For custom app: Can be an asset name uploaded to Discord Developer Portal
- **Editable:** Yes
- **See also:** `src/utilities/rpc/assets.py` for more image options

### discord_application_logo_small
- **Type:** String (URL or asset name)
- **Default:** `"https://raw.githubusercontent.com/6uhrmittag/Synth-Riders-DiscordRPC/refs/heads/master/assets/game_synthriders_logo.png"`
- **Description:** The small logo shown in Discord Rich Presence (overlay icon).
- **Format:** Same as `discord_application_logo_large`
- **Editable:** Yes

### synthriders_websocket_host
- **Type:** String (hostname or IP)
- **Default:** `"localhost"`
- **Description:** The hostname where the SynthRiders-Websockets-Mod is running.
- **Editable:** Yes
- **Use cases:**
  - `"localhost"` - Game and Discord RPC on same computer (default)
  - IP address - For 2-PC streaming setups where game runs on a different computer

### synthriders_websocket_port
- **Type:** String (port number)
- **Default:** `"9000"`
- **Description:** The port where the SynthRiders-Websockets-Mod is listening.
- **Editable:** Yes
- **Note:** Must match the port configured in SynthRiders mod's `MelonPreferences.cfg`

### image_upload_url
- **Type:** String (URL)
- **Default:** `"https://uguu.se/upload"`
- **Description:** The URL endpoint for uploading album artwork.
- **Editable:** Yes
- **Compatible services:** Any [uguu](https://github.com/topics/uguu) or [pomf-based](https://github.com/topics/pomf) file hosting service
- **Note:** Album art is temporary (deleted after 3 hours) and typically 100-500KB per upload

## Common Customizations

### Disable Promotion Button

Edit `config.json` and set:
```json
"promote_preference": false
```

### Use Custom Discord Application

1. Create a Discord application at https://discord.com/developers/applications
2. Upload your custom images as assets
3. Edit `config.json`:
```json
"discord_application_id": "YOUR_APP_ID",
"discord_application_logo_large": "your_large_asset_name",
"discord_application_logo_small": "your_small_asset_name"
```

### Change Websocket Port (2-PC Setup)

If running Synth Riders on a different PC:

1. On the game PC, edit SynthRiders' `MelonPreferences.cfg` to set `Host` to `0.0.0.0`
2. On the Discord PC, edit `config.json`:
```json
"synthriders_websocket_host": "192.168.1.XXX",
"synthriders_websocket_port": "9000"
```

### Use Alternative File Upload Service

Edit `config.json`:
```json
"image_upload_url": "https://alternative-service.com/upload"
```

Service must be compatible with uguu/pomf API format.

### Disable Automatic Startup

Edit `config.json`:
```json
"startup_preference": false
```

Then remove the application from Windows startup:
- Open Task Manager
- Go to Startup tab
- Disable "Synth Riders DiscordRPC"

## Troubleshooting Configuration Issues

### Changes Not Taking Effect

- Restart the DiscordRPC application after editing `config.json`
- Check the log file at: `%localappdata%\Synth Riders DiscordRPC\logs\log.txt`
- Ensure JSON syntax is valid (no trailing commas, proper quotes)

### Cannot Find config.json

If the file doesn't exist, re-run the setup executable to recreate it.

### Invalid Configuration Error

If the application reports invalid configuration:
1. Check that `rich_presence_install_location` matches the actual install location
2. Delete `config.json` and re-run setup
3. Check log file for specific error messages

### Album Art Not Showing

- Verify `image_upload_url` is accessible
- Check internet connection
- Some file hosts may be temporarily down - try an alternative service

## Advanced Configuration

### Image Assets

The application includes several logo variations in the `assets/` directory:
- `game_synthriders_logo.png` - Official logo
- `game_synthriders_logo_alt.png` - Alternative logo style
- `game_synthriders_logo_square.jpg` - Square version

You can host these yourself or use the GitHub URLs as shown in the default configuration.

### Timing Constants

To adjust timing behavior, edit `src/constants.py` (requires rebuilding the application):
- `PRESENCE_UPDATE_INTERVAL` - How often to update Discord (default: 15s)
- `DISCORD_RECONNECT_INTERVAL` - Discord reconnection delay (default: 15s)
- `WEBSOCKET_RECONNECT_INTERVAL` - Websocket reconnection delay (default: 5s)
- `GAME_CHECK_INTERVAL` - Game process check interval (default: 5s)

## Related Files

- **Configuration:** `config.json` (this document)
- **Images:** `src/utilities/rpc/assets.py`
- **Timing:** `src/constants.py`
- **Application Settings:** `config.py` (system-wide defaults)

## Support

If you have configuration issues not covered here:
1. Check the logs at `%localappdata%\Synth Riders DiscordRPC\logs\log.txt`
2. Open an issue at https://github.com/6uhrmittag/Synth-Riders-DiscordRPC/issues

