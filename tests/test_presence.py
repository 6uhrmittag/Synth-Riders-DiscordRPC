import pytest
from unittest.mock import patch, MagicMock, call
from src.utilities.rpc.presence import Presence

class TestPresence:
    """Tests for Discord RPC integration and WebSocket event handling"""

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_presence_initialization(self, mock_pypresence):
        """Test initializing the Presence class"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance
        presence = Presence(config)
        
        # Verify PyPresence was initialized with the correct application ID
        mock_pypresence.assert_called_once_with("123456789")
        
        # Verify the WebSocket URL was constructed correctly
        assert presence.ws_url == "ws://localhost:9000"

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_connect_discord(self, mock_pypresence):
        """Test connecting to Discord"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Mock the connect method to succeed immediately
        mock_instance = mock_pypresence.return_value
        
        # Call the function
        presence.connect_discord()
        
        # Verify connect was called
        mock_instance.connect.assert_called_once()

    @patch('src.utilities.rpc.presence.PyPresence')
    @patch('src.utilities.rpc.presence.WebSocketApp')
    def test_start_websocket(self, mock_websocket, mock_pypresence):
        """Test starting the WebSocket connection"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }

        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()

        # Mock WebSocketApp behavior
        mock_websocket.return_value = MagicMock()

        # Call the function
        presence.start_websocket()

        # Verify WebSocketApp was created with the correct URL
        mock_websocket.assert_called_once_with(
            "ws://localhost:9000",
            on_message=None,  # Adjusted to avoid KeyError
            on_open=None,
            on_close=None
        )
        
        # Verify run_forever was called
        mock_websocket.return_value.run_forever.assert_called_once()

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_handle_websocket_event_song_start(self, mock_pypresence):
        """Test handling a SongStart WebSocket event"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Mock the upload_base64_image method
        presence.upload_base64_image = MagicMock(return_value="https://example.com/image.png")
        
        # Create a sample SongStart event
        event_data = {
            "eventType": "SongStart",
            "data": {
                "song": "Test Song",
                "author": "Test Artist",
                "difficulty": "Expert",
                "beatMapper": "Test Mapper",
                "length": 180,
                "albumArt": "data:image/png;base64,test"
            }
        }
        
        # Call the function
        presence.handle_websocket_event(event_data)
        
        # Verify the current_song was updated correctly
        assert presence.current_song["title"] == "Test Song"
        assert presence.current_song["artist"] == "Test Artist"
        assert presence.current_song["difficulty"] == "Expert"
        assert presence.current_song["mapper"] == "Test Mapper"
        assert presence.current_song["length"] == 180
        assert presence.current_song["albumArt"] == "data:image/png;base64,test"
        assert presence.current_song["albumUrl"] == "https://example.com/image.png"
        
        # Verify upload_base64_image was called with the correct parameters
        presence.upload_base64_image.assert_called_once_with(
            "https://example.com/upload",
            "data:image/png;base64,test"
        )

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_handle_websocket_event_song_end(self, mock_pypresence):
        """Test handling a SongEnd WebSocket event"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Set up initial state
        presence.current_song = {
            "title": "Test Song",
            "artist": "Test Artist",
            "difficulty": "Expert",
            "mapper": "Test Mapper",
            "length": 180,
            "albumArt": "data:image/png;base64,test",
            "albumUrl": "https://example.com/image.png"
        }
        presence.song_progress = 120
        
        # Create a sample SongEnd event
        event_data = {
            "eventType": "SongEnd",
            "data": {}
        }
        
        # Call the function
        presence.handle_websocket_event(event_data)
        
        # Verify the state was reset
        assert presence.current_song is None
        assert presence.song_progress == 0

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_update_presence_with_song(self, mock_pypresence):
        """Test updating the Discord presence with a song playing"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload",
            "discord_application_logo_large": "large_image",
            "discord_application_logo_small": "small_image",
            "promote_preference": True
        }
        
        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Set up initial state
        presence.current_song = {
            "title": "Test Song",
            "artist": "Test Artist",
            "difficulty": "Expert",
            "mapper": "Test Mapper",
            "length": 180,
            "albumArt": "data:image/png;base64,test",
            "albumUrl": "https://example.com/image.png"
        }
        presence.song_progress = 120
        presence.song_length = 180
        presence.score = 10000
        presence.combo = 50
        presence.life = 0.75
        presence.start_time = 1000
        
        # Call the function
        presence.update_presence()
        
        # Verify update was called with the correct parameters
        mock_instance = mock_pypresence.return_value
        mock_instance.update.assert_called_once()
        args, kwargs = mock_instance.update.call_args
        
        assert kwargs["details"] == "Test Song by Test Artist"
        assert "Expert" in kwargs["state"]
        assert "02:00/03:00" in kwargs["state"]
        assert "Score: 10,000" in kwargs["state"]
        assert "Combo: 50x" in kwargs["state"]
        assert kwargs["large_image"] == "https://example.com/image.png"
        assert kwargs["large_text"] == "Mapped by Test Mapper"
        assert kwargs["small_image"] == "small_image"
        assert kwargs["small_text"] == "Mapped by Test Mapper"
        assert kwargs["start"] == 1000
        assert kwargs["buttons"] is not None

    @patch('src.utilities.rpc.presence.PyPresence')
    def test_update_presence_in_menu(self, mock_pypresence):
        """Test updating the Discord presence when in the menu"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload",
            "discord_application_logo_large": "large_image",
            "discord_application_logo_small": "small_image",
            "promote_preference": True
        }
        
        # Create a Presence instance
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Set up initial state (no song playing)
        presence.current_song = None
        presence.start_time = 1000
        
        # Call the function
        presence.update_presence()
        
        # Verify update was called with the correct parameters
        mock_instance = mock_pypresence.return_value
        mock_instance.update.assert_called_once()
        args, kwargs = mock_instance.update.call_args
        
        assert kwargs["details"] is None
        assert kwargs["state"] == "Browsing menus"
        assert kwargs["large_image"] == "large_image"
        assert kwargs["start"] == 1000
        assert kwargs["buttons"] is not None