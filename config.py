from os.path import join
from os import getenv


class Config:
    VERSION = "1.1.1"
    MAIN_EXECUTABLE_NAME = "Synth Riders DiscordRPC.exe"
    UNINSTALL_EXECUTABLE_NAME = "Uninstall Synth Riders DiscordRPC.exe"
    # APPLICATION_ID = "1342397301687189544" # Custom App for this project; can use custom images uploaded to the App
    APPLICATION_ID = "1124356298578870333" # Official Synth Riders SteamVR App; requires the images to be URLs
    SYNTH_RIDERS_PROCESS_NAME = "SynthRiders.exe"
    WEBSOCKET_HOST = "localhost"
    WEBSOCKET_PORT = "9000"
    IMAGE_UPLOAD_URL = "https://uguu.se/upload"