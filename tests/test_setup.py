import os
import tempfile
import pytest
from unittest.mock import patch, MagicMock
from src.bin.setup import (
    create_config_folder,
    write_config_to_file,
    get_config
)

class TestSetup:
    """Tests for setup logic"""

    def test_create_config_folder(self):
        """Test creating the config folder"""
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Mock console
            mock_console = MagicMock()
            
            # Config with the temporary directory
            config = {
                "rich_presence_install_location": temp_dir
            }
            
            # Call the function
            create_config_folder(mock_console, config)
            
            # Verify the config folder was created
            config_folder_path = os.path.join(temp_dir, "config")
            assert os.path.exists(config_folder_path)
            assert os.path.isdir(config_folder_path)
            
            # Verify console output
            mock_console.print.assert_called()

    def test_write_config_to_file(self):
        """Test writing the config to a file"""
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create config folder
            config_folder_path = os.path.join(temp_dir, "config")
            os.makedirs(config_folder_path, exist_ok=True)
            
            # Mock console
            mock_console = MagicMock()
            
            # Sample config
            config = {
                "rich_presence_install_location": temp_dir,
                "version": "1.1.1",
                "synthriders_install_location": "C:\\Games\\SynthRiders",
                "startup_preference": True,
                "keep_running_preference": True,
                "shortcut_preference": False,
                "promote_preference": True
            }
            
            # Call the function
            write_config_to_file(mock_console, config)
            
            # Verify the config file was created
            config_file_path = os.path.join(config_folder_path, "config.json")
            assert os.path.exists(config_file_path)
            assert os.path.isfile(config_file_path)
            
            # Verify console output
            mock_console.print.assert_called()
            
            # Read the config file and verify its contents
            import json
            with open(config_file_path, "r") as f:
                loaded_config = json.load(f)
            
            assert loaded_config == config

    @patch('src.bin.setup.get_synthriders_install_location')
    @patch('src.bin.setup.get_rich_presence_install_location')
    @patch('src.bin.setup.get_startup_preference')
    @patch('src.bin.setup.get_keep_running_preference')
    @patch('src.bin.setup.get_shortcut_preference')
    @patch('src.bin.setup.get_promote_preference')
    def test_get_config(self, mock_promote, mock_shortcut, mock_keep_running, 
                       mock_startup, mock_rich_presence, mock_synthriders):
        """Test getting the config from user input"""
        # Mock return values
        mock_synthriders.return_value = "C:\\Games\\SynthRiders"
        mock_rich_presence.return_value = "C:\\Apps\\SynthRPC"
        mock_startup.return_value = True
        mock_keep_running.return_value = True
        mock_shortcut.return_value = False
        mock_promote.return_value = True
        
        # Mock console
        mock_console = MagicMock()
        
        # Mock get_input to return the mocked values
        with patch('src.bin.setup.get_input', side_effect=[
            "C:\\Games\\SynthRiders",
            "C:\\Apps\\SynthRPC",
            True,
            True,
            False,
            True
        ]):
            # Call the function
            config = get_config(mock_console)
        
        # Verify the config has the expected values
        from config import Config
        from src.utilities.rpc.assets import DiscordAssets
        
        assert config["version"] == Config.VERSION
        assert config["synthriders_install_location"] == "C:\\Games\\SynthRiders"
        assert config["rich_presence_install_location"] == "C:\\Apps\\SynthRPC"
        assert config["startup_preference"] is True
        assert config["keep_running_preference"] is True
        assert config["shortcut_preference"] is False
        assert config["promote_preference"] is True
        assert config["discord_application_id"] == Config.APPLICATION_ID
        assert config["discord_application_logo_large"] == DiscordAssets.LARGE_IMAGE
        assert config["discord_application_logo_small"] == DiscordAssets.SMALL_IMAGE
        assert config["synthriders_websocket_host"] == Config.WEBSOCKET_HOST
        assert config["synthriders_websocket_port"] == Config.WEBSOCKET_PORT
        assert config["image_upload_url"] == Config.IMAGE_UPLOAD_URL

def test_dummy():
    """Basic test to ensure setup module is accessible."""
    assert True
