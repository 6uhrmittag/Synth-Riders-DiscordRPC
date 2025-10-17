"""
Tests for the Logger class
"""

import os
import tempfile
import pytest
from src.utilities.rpc.logger import Logger


class TestLogger:
    """Test suite for Logger functionality"""
    
    @pytest.fixture
    def temp_log_folder(self):
        """Create a temporary folder for log files"""
        with tempfile.TemporaryDirectory() as temp_dir:
            yield temp_dir
    
    @pytest.fixture
    def logger(self, temp_log_folder):
        """Create a logger instance with temporary log folder"""
        return Logger(log_folder=temp_log_folder)
    
    def test_logger_creates_log_folder(self, temp_log_folder):
        """Test that logger creates the log folder if it doesn't exist"""
        log_folder = os.path.join(temp_log_folder, "new_logs")
        logger = Logger(log_folder=log_folder)
        
        assert os.path.exists(log_folder)
        
        # Log file is created on first write
        logger.info("Test")
        assert os.path.exists(logger.log_file_path)
    
    def test_logger_info(self, logger):
        """Test logging info messages"""
        logger.info("Test info message")
        
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            assert "[INFO]" in content
            assert "Test info message" in content
    
    def test_logger_error(self, logger):
        """Test logging error messages"""
        logger.error("Test error message")
        
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            assert "[ERROR]" in content
            assert "Test error message" in content
    
    def test_logger_warning(self, logger):
        """Test logging warning messages"""
        logger.warning("Test warning message")
        
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            assert "[WARNING]" in content
            assert "Test warning message" in content
    
    def test_logger_clear(self, logger):
        """Test that clear() removes all log content"""
        # Write some log entries
        logger.info("First message")
        logger.error("Second message")
        
        # Verify content exists
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            assert len(content) > 0
        
        # Clear the log
        logger.clear()
        
        # Verify content is empty
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            assert len(content) == 0
    
    def test_logger_multiple_entries(self, logger):
        """Test multiple log entries are all written"""
        logger.info("Info 1")
        logger.warning("Warning 1")
        logger.error("Error 1")
        logger.info("Info 2")
        
        with open(logger.log_file_path, "r") as f:
            lines = f.readlines()
            assert len(lines) == 4
            assert any("Info 1" in line for line in lines)
            assert any("Warning 1" in line for line in lines)
            assert any("Error 1" in line for line in lines)
            assert any("Info 2" in line for line in lines)
    
    def test_logger_timestamp_format(self, logger):
        """Test that log entries include timestamps"""
        logger.info("Test message")
        
        with open(logger.log_file_path, "r") as f:
            content = f.read()
            # Should contain date/time information
            assert "[INFO]" in content
            assert "[" in content
            assert "]" in content

