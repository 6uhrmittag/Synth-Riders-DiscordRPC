"""
Application constants for timing and intervals.

These constants centralize commonly used values to make them easier to
tune and maintain. Changing values here affects application behavior
throughout the codebase.
"""

# Discord Rich Presence update interval
# How often to update Discord with current game state
PRESENCE_UPDATE_INTERVAL = 15  # seconds

# Discord connection retry interval
# How long to wait before retrying Discord connection if it fails
DISCORD_RECONNECT_INTERVAL = 15  # seconds

# WebSocket reconnection interval
# How long to wait before retrying websocket connection if it closes
WEBSOCKET_RECONNECT_INTERVAL = 5  # seconds

# Game process check interval
# How often to check if the game is still running when keep_running_preference is enabled
GAME_CHECK_INTERVAL = 5  # seconds

