import os
import json
import tempfile
import pytest
from config import Config

class TestConfig:
    """Tests for configuration handling"""

    def test_config_constants(self):
        """Test that the Config class has the expected constants"""
        assert hasattr(Config, 'VERSION')
        assert hasattr(Config, 'MAIN_EXECUTABLE_NAME')
        assert hasattr(Config, 'UNINSTALL_EXECUTABLE_NAME')
        assert hasattr(Config, 'APPLICATION_ID')
        assert hasattr(Config, 'SYNTH_RIDERS_PROCESS_NAME')
        assert hasattr(Config, 'WEBSOCKET_HOST')
        assert hasattr(Config, 'WEBSOCKET_PORT')
        assert hasattr(Config, 'IMAGE_UPLOAD_URL')

    def test_config_file_write_read(self):
        """Test writing and reading a config file using production config"""
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "config.json")
            
            # Use production config values
            config_data = {
                "version": Config.VERSION,
                "synthriders_install_location": "C:\\Games\\SynthRiders",
                "rich_presence_install_location": "C:\\Apps\\SynthRPC",
                "startup_preference": True,
                "keep_running_preference": True,
                "shortcut_preference": False,
                "promote_preference": True,
                "discord_application_id": Config.APPLICATION_ID,
                "synthriders_websocket_host": Config.WEBSOCKET_HOST,
                "synthriders_websocket_port": Config.WEBSOCKET_PORT,
                "image_upload_url": Config.IMAGE_UPLOAD_URL
            }
            
            # Write config to file
            with open(config_path, "w") as f:
                json.dump(config_data, f, indent=4)
            
            # Read config from file
            with open(config_path, "r") as f:
                loaded_config = json.load(f)
            
            # Verify the loaded config matches the original
            assert loaded_config == config_data
            assert loaded_config["version"] == Config.VERSION
            assert loaded_config["discord_application_id"] == Config.APPLICATION_ID
            assert loaded_config["synthriders_websocket_host"] == Config.WEBSOCKET_HOST
            assert loaded_config["synthriders_websocket_port"] == Config.WEBSOCKET_PORT
            assert loaded_config["image_upload_url"] == Config.IMAGE_UPLOAD_URL

    def test_config_version_migration(self):
        """Test migrating from an older config version"""
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = os.path.join(temp_dir, "config.json")
            
            # Sample old config data with a different version
            old_config = {
                "version": "1.0.0",  # Old version
                "synthriders_install_location": "C:\\Games\\SynthRiders",
                "rich_presence_install_location": "C:\\Apps\\SynthRPC",
                "startup_preference": True,
                "keep_running_preference": True,
                "shortcut_preference": False,
                "promote_preference": True,
                "discord_application_id": "old_app_id",  # Old app ID
                # Missing new fields
            }
            
            # Write old config to file
            with open(config_path, "w") as f:
                json.dump(old_config, f, indent=4)
            
            # Simulate migration by updating the config
            with open(config_path, "r") as f:
                config = json.load(f)
            
            # Update version and add missing fields
            config["version"] = Config.VERSION
            config["discord_application_id"] = Config.APPLICATION_ID
            if "synthriders_websocket_host" not in config:
                config["synthriders_websocket_host"] = Config.WEBSOCKET_HOST
            if "synthriders_websocket_port" not in config:
                config["synthriders_websocket_port"] = Config.WEBSOCKET_PORT
            if "image_upload_url" not in config:
                config["image_upload_url"] = Config.IMAGE_UPLOAD_URL
            
            # Write updated config to file
            with open(config_path, "w") as f:
                json.dump(config, f, indent=4)
            
            # Read the migrated config
            with open(config_path, "r") as f:
                migrated_config = json.load(f)
            
            # Verify the migration was successful
            assert migrated_config["version"] == Config.VERSION
            assert migrated_config["discord_application_id"] == Config.APPLICATION_ID
            assert migrated_config["synthriders_websocket_host"] == Config.WEBSOCKET_HOST
            assert migrated_config["synthriders_websocket_port"] == Config.WEBSOCKET_PORT
            assert migrated_config["image_upload_url"] == Config.IMAGE_UPLOAD_URL