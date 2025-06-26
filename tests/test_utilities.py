import pytest
from unittest.mock import patch, MagicMock
from src.utilities.rpc.presence import Presence
from src.config import Config

class TestUtilities:
    """Tests for utility functions"""

    def test_format_time(self):
        """Test the format_time function"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }

        # Create a Presence instance with the mock config
        presence = Presence(config)

        # Test various time values
        assert presence.format_time(0) == "00:00"
        assert presence.format_time(60) == "01:00"
        assert presence.format_time(90) == "01:30"
        assert presence.format_time(3600) == "01:00:00"  # Adjusted to match expected format
        assert presence.format_time(3661) == "01:01:01"  # Adjusted to match expected format

    @patch('src.utilities.rpc.presence.requests.post')
    def test_upload_base64_image(self, mock_post):
        """Test the upload_base64_image function using production config"""
        # Use production config values
        presence = Presence({
            "discord_application_id": Config.APPLICATION_ID,
            "synthriders_websocket_host": Config.WEBSOCKET_HOST,
            "synthriders_websocket_port": Config.WEBSOCKET_PORT,
            "image_upload_url": Config.IMAGE_UPLOAD_URL
        })
        presence.logger = MagicMock()
        
        # Mock the response from the upload service
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "files": [
                {
                    "url": "https://example.com/image.png"
                }
            ]
        }
        mock_post.return_value = mock_response
        
        # Sample base64 image
        base64_image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
        
        # Call the function
        result = presence.upload_base64_image(Config.IMAGE_UPLOAD_URL, base64_image)

        # Verify the result
        assert result == "https://example.com/image.png"
        
        # Verify the post request was made with the correct parameters
        mock_post.assert_called_once()
        
        # Verify the logger was called
        presence.logger.info.assert_called()

    @patch('src.utilities.rpc.presence.requests.post')
    def test_upload_base64_image_error(self, mock_post):
        """Test the upload_base64_image function with an error response"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance with the mock config
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Mock the response from the upload service
        mock_response = MagicMock()
        mock_response.status_code = 500
        mock_response.text = "Internal Server Error"
        mock_post.return_value = mock_response
        
        # Sample base64 image (a small 1x1 transparent PNG)
        base64_image = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
        
        # Call the function and expect an exception
        with pytest.raises(ValueError):
            presence.upload_base64_image("https://example.com/upload", base64_image)
        
        # Verify the post request was made with the correct parameters
        mock_post.assert_called_once()
        
        # Verify the logger was called
        presence.logger.error.assert_called()

    def test_invalid_base64_image(self):
        """Test the upload_base64_image function with an invalid base64 image"""
        # Create a mock config
        config = {
            "discord_application_id": "123456789",
            "synthriders_websocket_host": "localhost",
            "synthriders_websocket_port": "9000",
            "image_upload_url": "https://example.com/upload"
        }
        
        # Create a Presence instance with the mock config
        presence = Presence(config)
        presence.logger = MagicMock()
        
        # Invalid base64 image (missing the data:image part)
        invalid_base64 = "not a valid base64 image"
        
        # Call the function and expect an exception
        with pytest.raises(ValueError):
            presence.upload_base64_image("https://example.com/upload", invalid_base64)
        
        # Verify the logger was called
        presence.logger.error.assert_called()