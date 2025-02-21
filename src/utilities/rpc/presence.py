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

    def __init__(self, config: dict) -> None:
        self.config = config
        self.logger = Logger()


        self.presence = PyPresence(Config.APPLICATION_ID)
        self.ws_url = f"ws://{config.get('websocket_host', 'localhost')}:{config.get('websocket_port', 9000)}"

    def start(self) -> None:
        """
        Start the RPC
        """
        try:
            self.logger.clear()
            self.connect_discord()
            self.start_websocket()
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
                    "length": event_data.get("length", 0)
                }
                self.song_length = self.current_song["length"]
                self.song_progress = 0
                self.score = 0
                self.combo = 0
                self.life = 1.0

            elif event_type == "SongEnd":
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
            "label": "Want this status?",
            "url": "https://github.com/your/repo"
        }] if self.config.get("promote_preference") else None

        with self.lock:
            if self.current_song:
                time_str = self.format_time(self.song_progress)
                length_str = self.format_time(self.song_length)

                details = f"In Synth Riders: {self.current_song['title']} by {self.current_song['artist']}"
                state = (f"{self.current_song['difficulty']} | "
                        f"{time_str}/{length_str} | "
                        f"Score: {self.score:,} | "
                        f"Combo: {self.combo}x")

                self.presence.update(
                    details=details,
                    state=state,
                    large_image=DiscordAssets.LARGE_IMAGE,
                    large_text=f"Playing Synth Riders VR",
                    # large_text=f"Playing Synth RidersMapped by {self.current_song['mapper']}",
                    small_image=DiscordAssets.SMALL_IMAGE,
                    small_text=f"Mapped by {self.current_song['mapper']}",
                    # small_text=f"Life: {self.life*100:.0f}%",
                    buttons=buttons
                )
            else:
                self.presence.update(
                    details="Playing Synth Riders VR",
                    state="In menus",
                    large_image=DiscordAssets.LARGE_IMAGE,
                    buttons=buttons
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