"""
Mock WebSocket Server for SynthRiders-Websockets-Mod
Simulates the official SynthRiders-Websockets-Mod events for testing.

Based on: https://github.com/bookdude13/SynthRiders-Websockets-Mod

Usage:
    Run as standalone: python tests/mock_websocket_server.py
    Or import and use in tests
"""

import json
import time
import base64
from websocket_server import WebsocketServer

# Sample base64 encoded PNG (1x1 transparent pixel) for album art testing
SAMPLE_ALBUM_ART = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="


def create_song_start_event():
    """Create a SongStart event with realistic data"""
    return {
        "eventType": "SongStart",
        "data": {
            "song": "2 Phut Hon (Kaiz Remix)",
            "difficulty": "Master",
            "author": "Phao",
            "beatMapper": "ICHDerHorst",
            "length": 191.857,
            "bpm": 128.0,
            "albumArt": SAMPLE_ALBUM_ART
        }
    }


def create_play_time_event(play_time_ms):
    """Create a PlayTime event"""
    return {
        "eventType": "PlayTime",
        "data": {
            "playTimeMS": play_time_ms
        }
    }


def create_note_hit_event(score, combo, play_time_ms):
    """Create a NoteHit event"""
    return {
        "eventType": "NoteHit",
        "data": {
            "score": score,
            "combo": combo,
            "multiplier": min(6, (combo // 10) + 1),
            "completed": combo,
            "lifeBarPercent": 1.0,
            "playTimeMS": play_time_ms
        }
    }


def create_note_miss_event(multiplier, life_bar_percent, play_time_ms):
    """Create a NoteMiss event"""
    return {
        "eventType": "NoteMiss",
        "data": {
            "multiplier": multiplier,
            "lifeBarPercent": life_bar_percent,
            "playTimeMS": play_time_ms
        }
    }


def create_song_end_event():
    """Create a SongEnd event"""
    return {
        "eventType": "SongEnd",
        "data": {
            "song": "2 Phut Hon (Kaiz Remix)",
            "perfect": 350,
            "normal": 126,
            "bad": 281,
            "fail": 2,
            "highestCombo": 482
        }
    }


def create_return_to_menu_event():
    """Create a ReturnToMenu event"""
    return {
        "eventType": "ReturnToMenu",
        "data": {}
    }


def create_scene_change_event(scene_name):
    """Create a SceneChange event"""
    return {
        "eventType": "SceneChange",
        "data": {
            "sceneName": scene_name
        }
    }


def create_enter_special_event():
    """Create an EnterSpecial event"""
    return {
        "eventType": "EnterSpecial",
        "data": {}
    }


def create_complete_special_event():
    """Create a CompleteSpecial event"""
    return {
        "eventType": "CompleteSpecial",
        "data": {}
    }


def create_fail_special_event():
    """Create a FailSpecial event"""
    return {
        "eventType": "FailSpecial",
        "data": {}
    }


class MockSynthRidersWebsocketServer:
    """Mock WebSocket server that simulates SynthRiders-Websockets-Mod"""
    
    def __init__(self, host='localhost', port=9000):
        self.host = host
        self.port = port
        self.server = None
        self.clients = []
    
    def new_client(self, client, server):
        """Called when a new client connects"""
        self.clients.append(client)
        print(f"New client connected: {client['id']}")
    
    def client_left(self, client, server):
        """Called when a client disconnects"""
        if client in self.clients:
            self.clients.remove(client)
        print(f"Client disconnected: {client['id']}")
    
    def send_event(self, event):
        """Send an event to all connected clients"""
        message = json.dumps(event)
        if self.server:
            self.server.send_message_to_all(message)
            print(f"Sent event: {event['eventType']}")
    
    def simulate_song_playthrough(self):
        """Simulate a complete song playthrough"""
        print("\n=== Simulating Song Playthrough ===")
        
        # Start the song
        self.send_event(create_song_start_event())
        time.sleep(2)
        
        # Simulate 10 seconds of gameplay with note hits
        for i in range(10):
            play_time_ms = (i + 1) * 1000
            self.send_event(create_play_time_event(play_time_ms))
            
            # Every second, send a note hit
            score = (i + 1) * 100
            combo = i + 1
            self.send_event(create_note_hit_event(score, combo, play_time_ms))
            
            time.sleep(1)
        
        # End the song
        self.send_event(create_song_end_event())
        time.sleep(1)
        
        # Scene change to game end
        self.send_event(create_scene_change_event("3.GameEnd"))
        print("=== Song Playthrough Complete ===\n")
    
    def simulate_return_to_menu(self):
        """Simulate returning to menu during a song"""
        print("\n=== Simulating Return to Menu ===")
        
        # Start a song
        self.send_event(create_song_start_event())
        time.sleep(2)
        
        # Play a bit
        self.send_event(create_play_time_event(5000))
        self.send_event(create_note_hit_event(500, 5, 5000))
        time.sleep(1)
        
        # Return to menu
        self.send_event(create_return_to_menu_event())
        print("=== Returned to Menu ===\n")
    
    def start(self):
        """Start the mock websocket server"""
        print(f"Starting mock SynthRiders WebSocket server on ws://{self.host}:{self.port}")
        self.server = WebsocketServer(host=self.host, port=self.port)
        self.server.set_fn_new_client(self.new_client)
        self.server.set_fn_client_left(self.client_left)
        
        print("Server started! Waiting for connections...")
        print("Press Ctrl+C to stop")
        
        try:
            self.server.run_forever(threaded=False)
        except KeyboardInterrupt:
            print("\nShutting down server...")
    
    def start_threaded(self):
        """Start the server in a separate thread (for testing)"""
        import threading
        self.server = WebsocketServer(host=self.host, port=self.port)
        self.server.set_fn_new_client(self.new_client)
        self.server.set_fn_client_left(self.client_left)
        
        thread = threading.Thread(target=self.server.run_forever, daemon=True)
        thread.start()
        time.sleep(0.5)  # Give server time to start
        return thread


def run_interactive_server():
    """Run the server in interactive mode with test scenarios"""
    server = MockSynthRidersWebsocketServer()
    
    # Start server in a thread
    import threading
    server_thread = threading.Thread(target=server.start, daemon=True)
    server_thread.start()
    
    # Wait for clients to connect
    time.sleep(2)
    
    print("\n" + "="*50)
    print("Mock Server Commands:")
    print("  1 - Simulate full song playthrough")
    print("  2 - Simulate return to menu")
    print("  3 - Send SongStart event")
    print("  4 - Send PlayTime event")
    print("  5 - Send NoteHit event")
    print("  6 - Send SongEnd event")
    print("  q - Quit")
    print("="*50 + "\n")
    
    try:
        while True:
            if not server.clients:
                print("Waiting for client connection...")
                time.sleep(2)
                continue
            
            command = input("Enter command: ").strip().lower()
            
            if command == '1':
                server.simulate_song_playthrough()
            elif command == '2':
                server.simulate_return_to_menu()
            elif command == '3':
                server.send_event(create_song_start_event())
            elif command == '4':
                server.send_event(create_play_time_event(10000))
            elif command == '5':
                server.send_event(create_note_hit_event(1000, 10, 10000))
            elif command == '6':
                server.send_event(create_song_end_event())
            elif command == 'q':
                break
            else:
                print("Unknown command")
    
    except KeyboardInterrupt:
        print("\nShutting down...")


if __name__ == "__main__":
    print("SynthRiders Mock WebSocket Server")
    print("=" * 50)
    run_interactive_server()

