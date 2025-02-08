import pytest
import io
import os
from excel_lib import file_to_io_stream, is_match

@pytest.fixture
def sample_file(tmp_path):
    """Creates a temporary test file"""
    file_path = tmp_path / "test_file.txt"
    with open(file_path, "wb") as f:
        f.write(b"Test file content")
    return file_path

def test_file_to_io_stream_valid_file(sample_file):
    """Checks if the function correctly returns a byte stream for an existing file"""
    stream = file_to_io_stream(sample_file)
    assert isinstance(stream, io.BytesIO)
    assert stream.getvalue() == b"Test file content"

def test_file_to_io_stream_non_existent():
    """Checks if the function raises an exception for a non-existent file"""
    with pytest.raises(FileNotFoundError):
        file_to_io_stream("non_existent_file.txt")

def test_is_match_identical_strings():
    """Checks if the function correctly compares identical strings"""
    assert is_match("Test", "Test") is True
    assert is_match("123", "123") is True

def test_is_match_whitespace_handling():
    """Checks if the function ignores whitespace"""
    assert is_match(" Test ", "Test") is True
    assert is_match("\tHello\n", "Hello") is True

def test_is_match_different_strings():
    """Checks if the function correctly detects different values"""
    assert is_match("Test", "test") is False
    assert is_match("123", "124") is False

def test_is_match_numbers():
    """Checks if numbers are correctly compared"""
    assert is_match(5, 5) is True
    assert is_match(5, 10) is False

def test_is_match_different_types():
    """Checks if different types return False"""
    assert is_match(5, "5") is False
    assert is_match(3.14, "3.14") is False
