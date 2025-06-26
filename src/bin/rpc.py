import sys
import threading
from os.path import exists, join, abspath, dirname, normcase, normpath
from json import loads
from src.utilities.rpc import Presence
from src.utilities.tray import SystemTray

config_path = join(abspath(dirname(sys.executable)), "config/config.json")

if not exists(config_path):
    raise Exception(f"Config file does not exist, {config_path}")

with open(config_path, "r") as f:
    config = loads(f.read())
    if normpath(normcase(config["rich_presence_install_location"])) != normpath(
        normcase(abspath(dirname(sys.executable)))
    ):
        raise Exception(
            "The rich presence install location in the config file does not match the actual install location. Please update the config file, or setup the RPC again"
        )

# Create presence instance
presence = Presence(config)

# Create system tray
tray = SystemTray(presence)

# Run presence in background thread
presence_thread = threading.Thread(target=presence.start, daemon=True)
presence_thread.start()

# Run system tray in main thread (required for GUI)
tray.run()
