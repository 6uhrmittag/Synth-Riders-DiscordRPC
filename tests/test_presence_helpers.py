"""
Tests for Presence helper methods
"""

import pytest
from unittest.mock import Mock, patch
from src.utilities.rpc.presence import Presence


class TestPresenceHelpers:
    """Test suite for Presence helper methods"""
    
    @pytest.fixture
    def mock_config(self):
        """Create a mock configuration"""
        return {
            "discord_application_id": "test_id",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "discord_application_logo_large": "large_logo",
            "discord_application_logo_small": "small_logo",
            "image_upload_url": "https://test.upload.com/upload",
            "keep_running_preference": False,
            "promote_preference": False
        }
    
    @pytest.fixture
    def presence(self, mock_config):
        """Create a Presence instance with mocked PyPresence"""
        with patch('src.utilities.rpc.presence.PyPresence'):
            presence = Presence(mock_config)
            return presence
    
    def test_format_time_zero_seconds(self, presence):
        """Test format_time with 0 seconds"""
        result = presence.format_time(0)
        assert result == "00:00"
    
    def test_format_time_under_one_minute(self, presence):
        """Test format_time with seconds under 1 minute"""
        result = presence.format_time(45)
        assert result == "00:45"
    
    def test_format_time_exact_minutes(self, presence):
        """Test format_time with exact minutes"""
        result = presence.format_time(120)  # 2 minutes
        assert result == "02:00"
    
    def test_format_time_minutes_and_seconds(self, presence):
        """Test format_time with minutes and seconds"""
        result = presence.format_time(191.857)  # 3 minutes 11.857 seconds
        assert result == "03:11"
    
    def test_format_time_large_value(self, presence):
        """Test format_time with large values"""
        result = presence.format_time(3599)  # 59:59
        assert result == "59:59"
    
    @patch('src.utilities.rpc.presence.pids')
    @patch('src.utilities.rpc.presence.Process')
    def test_synth_riders_process_exists_true(self, mock_process_class, mock_pids, presence):
        """Test synth_riders_process_exists when process is running"""
        # Mock pids to return [1234]
        mock_pids.return_value = [1234]
        
        # Mock Process to return a process with name SynthRiders.exe
        mock_process = Mock()
        mock_process.name.return_value = "SynthRiders.exe"
        mock_process_class.return_value = mock_process
        
        result = presence.synth_riders_process_exists()
        assert result is True
    
    @patch('src.utilities.rpc.presence.pids')
    @patch('src.utilities.rpc.presence.Process')
    def test_synth_riders_process_exists_false(self, mock_process_class, mock_pids, presence):
        """Test synth_riders_process_exists when process is not running"""
        # Mock pids to return [1234, 5678]
        mock_pids.return_value = [1234, 5678]
        
        # Mock Process to return processes with different names
        mock_process = Mock()
        mock_process.name.return_value = "OtherProcess.exe"
        mock_process_class.return_value = mock_process
        
        result = presence.synth_riders_process_exists()
        assert result is False
    
    @patch('src.utilities.rpc.presence.pids')
    @patch('src.utilities.rpc.presence.Process')
    def test_synth_riders_process_exists_handles_no_such_process(
        self, mock_process_class, mock_pids, presence
    ):
        """Test synth_riders_process_exists handles NoSuchProcess exception"""
        from psutil import NoSuchProcess
        
        # Mock pids to return [1234]
        mock_pids.return_value = [1234]
        
        # Mock Process to raise NoSuchProcess
        mock_process_class.side_effect = NoSuchProcess(1234)
        
        result = presence.synth_riders_process_exists()
        assert result is False
    
    def test_handle_websocket_event_song_start(self, presence):
        """Test handle_websocket_event with SongStart event"""
        event = {
            "eventType": "SongStart",
            "data": {
                "song": "Test Song",
                "author": "Test Artist",
                "difficulty": "Hard",
                "beatMapper": "Test Mapper",
                "length": 180.0,
                "albumArt": "data:image/png;base64,test"
            }
        }
        
        # Mock the image upload
        with patch.object(presence, 'upload_base64_image', return_value="https://test.url/image.png"):
            presence.handle_websocket_event(event)
        
        assert presence.current_song is not None
        assert presence.current_song["title"] == "Test Song"
        assert presence.current_song["artist"] == "Test Artist"
        assert presence.current_song["difficulty"] == "Hard"
        assert presence.current_song["mapper"] == "Test Mapper"
        assert presence.current_song["length"] == 180.0
        assert presence.song_length == 180.0
        assert presence.song_progress == 0
        assert presence.score == 0
        assert presence.combo == 0
        assert presence.life == 1.0
    
    def test_handle_websocket_event_play_time(self, presence):
        """Test handle_websocket_event with PlayTime event"""
        event = {
            "eventType": "PlayTime",
            "data": {
                "playTimeMS": 15000
            }
        }
        
        presence.handle_websocket_event(event)
        assert presence.song_progress == 15.0  # Converted from MS to seconds
    
    def test_handle_websocket_event_note_hit(self, presence):
        """Test handle_websocket_event with NoteHit event"""
        event = {
            "eventType": "NoteHit",
            "data": {
                "score": 1000,
                "combo": 15,
                "lifeBarPercent": 0.95
            }
        }
        
        presence.handle_websocket_event(event)
        assert presence.score == 1000
        assert presence.combo == 15
        assert presence.life == 0.95
    
    def test_handle_websocket_event_song_end(self, presence):
        """Test handle_websocket_event with SongEnd event"""
        # Set up a current song first
        presence.current_song = {"title": "Test"}
        
        event = {
            "eventType": "SongEnd",
            "data": {}
        }
        
        presence.handle_websocket_event(event)
        assert presence.current_song is None
        assert presence.song_progress == 0
    
    def test_handle_websocket_event_return_to_menu(self, presence):
        """Test handle_websocket_event with ReturnToMenu event"""
        # Set up a current song first
        presence.current_song = {"title": "Test"}
        
        event = {
            "eventType": "ReturnToMenu",
            "data": {}
        }
        
        presence.handle_websocket_event(event)
        assert presence.current_song is None
        assert presence.song_progress == 0
    
    def test_handle_websocket_event_scene_change_game_end(self, presence):
        """Test handle_websocket_event with SceneChange to GameEnd"""
        # Set up a current song first
        presence.current_song = {"title": "Test"}
        
        event = {
            "eventType": "SceneChange",
            "data": {
                "sceneName": "3.GameEnd"
            }
        }
        
        presence.handle_websocket_event(event)
        assert presence.current_song is None

