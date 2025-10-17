"""
Tests for WebSocket integration with mock server
"""

import pytest
import time
import json
from unittest.mock import Mock, patch, MagicMock
from tests.mock_websocket_server import (
    MockSynthRidersWebsocketServer,
    create_song_start_event,
    create_play_time_event,
    create_note_hit_event,
    create_song_end_event,
    create_return_to_menu_event,
    create_scene_change_event
)


class TestWebsocketIntegration:
    """Test suite for WebSocket integration"""
    
    @pytest.fixture
    def mock_server(self):
        """Create and start a mock websocket server"""
        server = MockSynthRidersWebsocketServer(host='localhost', port=9001)
        thread = server.start_threaded()
        yield server
        # Cleanup: server thread is daemon, so it will be killed automatically
    
    def test_mock_server_starts(self, mock_server):
        """Test that the mock server starts successfully"""
        assert mock_server.server is not None
        assert mock_server.host == 'localhost'
        assert mock_server.port == 9001
    
    def test_create_song_start_event(self):
        """Test SongStart event creation"""
        event = create_song_start_event()
        
        assert event["eventType"] == "SongStart"
        assert "song" in event["data"]
        assert "difficulty" in event["data"]
        assert "author" in event["data"]
        assert "beatMapper" in event["data"]
        assert "length" in event["data"]
        assert "bpm" in event["data"]
        assert "albumArt" in event["data"]
        assert event["data"]["albumArt"].startswith("data:image/png;base64,")
    
    def test_create_play_time_event(self):
        """Test PlayTime event creation"""
        event = create_play_time_event(15000)
        
        assert event["eventType"] == "PlayTime"
        assert event["data"]["playTimeMS"] == 15000
    
    def test_create_note_hit_event(self):
        """Test NoteHit event creation"""
        event = create_note_hit_event(score=1000, combo=10, play_time_ms=10000)
        
        assert event["eventType"] == "NoteHit"
        assert event["data"]["score"] == 1000
        assert event["data"]["combo"] == 10
        assert event["data"]["playTimeMS"] == 10000
        assert "multiplier" in event["data"]
        assert "completed" in event["data"]
        assert "lifeBarPercent" in event["data"]
    
    def test_create_song_end_event(self):
        """Test SongEnd event creation"""
        event = create_song_end_event()
        
        assert event["eventType"] == "SongEnd"
        assert "song" in event["data"]
        assert "perfect" in event["data"]
        assert "normal" in event["data"]
        assert "bad" in event["data"]
        assert "fail" in event["data"]
        assert "highestCombo" in event["data"]
    
    def test_create_return_to_menu_event(self):
        """Test ReturnToMenu event creation"""
        event = create_return_to_menu_event()
        
        assert event["eventType"] == "ReturnToMenu"
        assert event["data"] == {}
    
    def test_create_scene_change_event(self):
        """Test SceneChange event creation"""
        event = create_scene_change_event("3.GameEnd")
        
        assert event["eventType"] == "SceneChange"
        assert event["data"]["sceneName"] == "3.GameEnd"
    
    def test_event_json_serialization(self):
        """Test that all events can be JSON serialized"""
        events = [
            create_song_start_event(),
            create_play_time_event(1000),
            create_note_hit_event(500, 5, 5000),
            create_song_end_event(),
            create_return_to_menu_event(),
            create_scene_change_event("3.GameEnd")
        ]
        
        for event in events:
            # Should not raise an exception
            json_str = json.dumps(event)
            # Should be able to decode it back
            decoded = json.loads(json_str)
            assert decoded["eventType"] == event["eventType"]


class TestPresenceWithMockWebsocket:
    """Test Presence class with mocked websocket connection"""
    
    @pytest.fixture
    def mock_config(self):
        """Create a mock configuration"""
        return {
            "discord_application_id": "test_id",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9001",
            "discord_application_logo_large": "large_logo",
            "discord_application_logo_small": "small_logo",
            "image_upload_url": "https://test.upload.com/upload",
            "keep_running_preference": False,
            "promote_preference": False
        }
    
    @pytest.fixture
    def mock_presence(self, mock_config):
        """Create a Presence instance with all external dependencies mocked"""
        with patch('src.utilities.rpc.presence.PyPresence') as mock_pypresence, \
             patch('src.utilities.rpc.presence.WebSocketApp') as mock_ws_app:
            
            from src.utilities.rpc.presence import Presence
            
            # Mock PyPresence
            mock_pypresence_instance = MagicMock()
            mock_pypresence.return_value = mock_pypresence_instance
            
            # Create presence instance
            presence = Presence(mock_config)
            presence.presence = mock_pypresence_instance
            
            yield presence
    
    def test_presence_song_start_updates_state(self, mock_presence):
        """Test that SongStart event updates presence state correctly"""
        event = create_song_start_event()
        
        with patch.object(mock_presence, 'upload_base64_image', return_value="https://test.url/image.png"):
            mock_presence.handle_websocket_event(event)
        
        assert mock_presence.current_song is not None
        assert mock_presence.current_song["title"] == event["data"]["song"]
        assert mock_presence.current_song["artist"] == event["data"]["author"]
        assert mock_presence.current_song["difficulty"] == event["data"]["difficulty"]
        assert mock_presence.song_length == event["data"]["length"]
        assert mock_presence.song_progress == 0
    
    def test_presence_play_time_updates_progress(self, mock_presence):
        """Test that PlayTime event updates song progress"""
        event = create_play_time_event(30000)  # 30 seconds
        
        mock_presence.handle_websocket_event(event)
        
        assert mock_presence.song_progress == 30.0
    
    def test_presence_note_hit_updates_score(self, mock_presence):
        """Test that NoteHit event updates score and combo"""
        event = create_note_hit_event(score=2500, combo=25, play_time_ms=25000)
        
        mock_presence.handle_websocket_event(event)
        
        assert mock_presence.score == 2500
        assert mock_presence.combo == 25
        assert mock_presence.life == 1.0
    
    def test_presence_song_end_clears_state(self, mock_presence):
        """Test that SongEnd event clears current song"""
        # Set up initial state
        mock_presence.current_song = {"title": "Test Song"}
        mock_presence.song_progress = 100.0
        
        event = create_song_end_event()
        mock_presence.handle_websocket_event(event)
        
        assert mock_presence.current_song is None
        assert mock_presence.song_progress == 0
    
    def test_presence_return_to_menu_clears_state(self, mock_presence):
        """Test that ReturnToMenu event clears current song"""
        # Set up initial state
        mock_presence.current_song = {"title": "Test Song"}
        mock_presence.song_progress = 50.0
        
        event = create_return_to_menu_event()
        mock_presence.handle_websocket_event(event)
        
        assert mock_presence.current_song is None
        assert mock_presence.song_progress == 0
    
    def test_presence_scene_change_game_end_clears_state(self, mock_presence):
        """Test that SceneChange to GameEnd clears current song"""
        # Set up initial state
        mock_presence.current_song = {"title": "Test Song"}
        
        event = create_scene_change_event("3.GameEnd")
        mock_presence.handle_websocket_event(event)
        
        assert mock_presence.current_song is None
    
    def test_presence_update_with_song_playing(self, mock_presence):
        """Test update_presence when a song is playing"""
        # Set up a song playing state
        with patch.object(mock_presence, 'upload_base64_image', return_value="https://test.url/image.png"):
            mock_presence.handle_websocket_event(create_song_start_event())
        
        mock_presence.song_progress = 60.0
        mock_presence.score = 5000
        mock_presence.combo = 50
        
        # Update presence
        mock_presence.update_presence()
        
        # Verify presence.update was called
        assert mock_presence.presence.update.called
        call_args = mock_presence.presence.update.call_args
        
        # Verify the details include song info
        assert "2 Phut Hon" in call_args[1]["details"]
        assert "Phao" in call_args[1]["details"]
        assert "Master" in call_args[1]["state"]
        assert "5,000" in call_args[1]["state"]
        assert "50x" in call_args[1]["state"]
    
    def test_presence_update_without_song(self, mock_presence):
        """Test update_presence when no song is playing"""
        # Ensure no song is playing
        mock_presence.current_song = None
        
        # Update presence
        mock_presence.update_presence()
        
        # Verify presence.update was called
        assert mock_presence.presence.update.called
        call_args = mock_presence.presence.update.call_args
        
        # Verify it shows browsing menus
        assert call_args[1]["state"] == "Browsing menus"
        assert call_args[1]["details"] is None

