import sys
from os.path import exists, join, abspath, dirname, normcase, normpath
from json import loads
import os
import ctypes
import requests
import webbrowser
from src.utilities.rpc import Presence


def check_for_update(current_version: str) -> None:
    """Check GitHub for a new release and notify the user."""
    try:
        response = requests.get(
            "https://api.github.com/repos/6uhrmittag/Synth-Riders-DiscordRPC/releases/latest",
            timeout=10,
        )
        if response.status_code == 200:
            data = response.json()
            tag = data.get("tag_name") or data.get("name")
            release_url = data.get("html_url")
            if tag:
                latest = tag.lstrip("v")
                if latest != current_version:
                    message = (
                        f"A new version ({latest}) is available. Open download page?"
                    )
                    if os.name == "nt":
                        result = ctypes.windll.user32.MessageBoxW(
                            None,
                            message,
                            "Update Available",
                            1,
                        )
                        if result == 1 and release_url:
                            webbrowser.open(release_url)
                    else:
                        print(message)
                        if release_url:
                            webbrowser.open(release_url)
    except Exception:
        # Silently ignore update check errors
        pass

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

check_for_update(config.get("version", "0"))

presence = Presence(config)
presence.start()
