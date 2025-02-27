import os
import re
# from sqlite3 import Connection
from time import sleep, time

from psutil import NoSuchProcess, Process, pids
from pypresence import Presence as PyPresence

from config import Config
from src.utilities.rpc import (
    DiscordAssets,
    Logger,
)

# required for Synth Riders
import json
import threading
import time

from websocket import WebSocketApp

# required for image upload
import requests
import base64
import tempfile


class Presence:
    logger: Logger
    presence: PyPresence
    ws = None
    current_song = None
    song_progress = 0
    song_length = 0
    score = 0
    combo = 0
    life = 1.0
    lock = threading.Lock()
    connected = False
    start_time = 0

    def __init__(self, config: dict) -> None:
        self.config = config
        self.logger = Logger()


        self.presence = PyPresence(self.config.get("discord_application_id"))
        self.ws_url = f"ws://{self.config.get("synthriders_websocket_host")}:{self.config.get("synthriders_websocket_port")}"

    def start(self) -> None:
        """
        Start the RPC
        """
        try:
            self.logger.clear()
            self.connect_discord()
            self.start_websocket()

            self.start_time = int(time.time())
            self.rpc_loop()
        except Exception as e:
            self.logger.error(f"An error occurred: {e}")

    def connect_discord(self):
        while True:
            try:
                self.presence.connect()
                break
            except Exception as e:
                self.logger.info("Waiting for Discord...")
                sleep(15)

    def start_websocket(self):
        def on_message(ws, message):
            try:
                data = json.loads(message)
                self.handle_websocket_event(data)
            except Exception as e:
                self.logger.error(f"WebSocket error: {e}")

        def on_open(ws):
            self.logger.info("Connected to SynthRiders WebSocket")
            self.connected = True

        def on_close(ws, close_status_code, close_msg):
            self.logger.info("WebSocket connection closed")
            self.connected = False
            if self.synth_riders_process_exists():
                sleep(5)
                self.start_websocket()

        self.ws = WebSocketApp(self.ws_url,
                             on_message=on_message,
                             on_open=on_open,
                             on_close=on_close)

        ws_thread = threading.Thread(target=self.ws.run_forever)
        ws_thread.daemon = True
        ws_thread.start()

    def upload_base64_image(self, upload_url: str, base64_string: str) -> str:
        # Extract the base64 part from the data URL
        match = re.match(r'data:image/\w+;base64,(.*)', base64_string)
        if not match:
            self.logger.error("Invalid base64 image format")
            raise ValueError("Invalid base64 image format")

        image_data = base64.b64decode(match.group(1))

        # Create a temporary file to store the image
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
            temp_file.write(image_data)
            temp_file_path = temp_file.name

        # Upload
        url = self.config.get("image_upload_url")
        with open(temp_file_path, "rb") as image_file:
            response = requests.post(url, files={"files[]": image_file})
            self.logger.info(f"Upload response: {response.text}")

        # Remove the temporary file after upload
        try:
            os.remove(temp_file_path)
        except OSError:
            self.logger.error("Failed to remove temporary file")
            pass

        # Parse and return the URL from response
        if response.status_code == 200:
            try:
                self.logger.info(f"Upload response: {response.json()}")
                return response.json()["files"][0]["url"]
            except (KeyError, IndexError):
                self.logger.error(f"Unexpected response format: {response.text}")
                raise ValueError("Unexpected response format")
        else:
            self.logger.error(f"Unexpected response format: {response.text}")
            raise ValueError(f"Upload failed with status code {response.status_code}: {response.text}")


    def handle_websocket_event(self, data):
        event_type = data.get("eventType")
        event_data = data.get("data", {})

        with self.lock:
            if event_type == "SongStart":
                self.current_song = {
                    "title": event_data.get("song", "Unknown Song"),
                    "artist": event_data.get("author", "Unknown Artist"),
                    "difficulty": event_data.get("difficulty", "Unknown"),
                    "mapper": event_data.get("beatMapper", "Unknown Mapper"),
                    "length": event_data.get("length", 0),
                    "albumArt": event_data.get("albumArt", None)
                }
                self.song_length = self.current_song["length"]
                self.song_progress = 0
                self.score = 0
                self.combo = 0
                self.life = 1.0

                self.logger.info(f"Current song data: {self.current_song}")

                # upload albumArt
                self.current_song["albumUrl"] = self.upload_base64_image(event_data.get("albumArt")) or None


            elif event_type == "SongEnd" or event_type == "ReturnToMenu":
                self.current_song = None
                self.song_progress = 0

            elif event_type == "PlayTime":
                self.song_progress = event_data.get("playTimeMS", 0) / 1000

            elif event_type == "NoteHit":
                self.score = event_data.get("score", 0)
                self.combo = event_data.get("combo", 0)
                self.life = event_data.get("lifeBarPercent", 1.0)

            elif event_type == "SceneChange":
                if event_data.get("sceneName") == "3.GameEnd":
                    self.current_song = None




    def rpc_loop(self):
        """
        Loop to keep the RPC running
        """
        while True:
            if not self.synth_riders_process_exists():
                self.handle_game_exit()
                break

            self.update_presence()
            sleep(15)

    def update_presence(self):
        buttons = [{
            "label": "Want this status too?",
            "url": "https://github.com/6uhrmittag/Synth-Riders-DiscordRPC"
        }] if self.config.get("promote_preference") else None

        with self.lock:
            if self.current_song:
                time_str = self.format_time(self.song_progress)
                length_str = self.format_time(self.song_length)

                details = f"{self.current_song['title']} by {self.current_song['artist']}"
                state = (f"{self.current_song['difficulty']} | "
                        f"{time_str}/{length_str} | "
                        f"Score: {self.score:,} | "
                        f"Combo: {self.combo}x")

                self.presence.update(
                    details=details,
                    state=state,
                    large_image=self.current_song['albumUrl'] or self.config.get("discord_application_logo_large"),
                    #large_text=f"Playing Synth Riders VR",
                    large_text=f"Mapped by {self.current_song['mapper']}",
                    small_image=self.config.get("discord_application_logo_small"),
                    small_text=f"Mapped by {self.current_song['mapper']}",
                    # small_text=f"Life: {self.life*100:.0f}%",
                    buttons=buttons,
                    start=self.start_time
                )
            else:
                self.presence.update(
                    details=None,
                    state="Browsing menus",
                    large_image=self.config.get("discord_application_logo_large"),
                    buttons=buttons,
                    start=self.start_time
                )

    def format_time(self, seconds):
        return time.strftime("%M:%S", time.gmtime(seconds))

    def handle_game_exit(self):
        self.logger.info("Synth Riders closed")
        self.presence.clear()
        if self.config.get("keep_running_preference"):
            while not self.synth_riders_process_exists():
                sleep(5)
            self.start()

    def synth_riders_process_exists(self):
        """
        Check whether the Wuthering Waves process is running

        :return: True if the process is running, False otherwise
        """
        for pid in pids():
            try:
                if Process(pid).name() == Config.SYNTH_RIDERS_PROCESS_NAME:
                    return True
            except NoSuchProcess:
                pass
        return False