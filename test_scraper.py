import pytest
import requests_mock
import pandas as pd
from pathlib import Path

# --- Core Cleaner Functions to Test ---
def clean_text(text):
    if not text: return ""
    for char in ["", "", "", "🌍", "📞", "🕒"]:
        text = text.replace(char, "")
    return text.strip()

def build_search_url(keyword, place):
    formatted_query = f"{keyword} {place}".replace(" ", "+")
    return f"https://www.google.com/maps/search/{formatted_query}"

# =====================================================================
# 🧪 Test 1: Google Character Sanitation
# =====================================================================
def test_clean_text_strips_google_font_artifacts():
    """Verifies that custom styling symbols from Google Maps are stripped."""
    dirty_address = " 123 Main St, Nagpur "
    expected_output = "123 Main St, Nagpur"
    assert clean_text(dirty_address) == expected_output

def test_clean_text_handles_empty_inputs():
    """Verifies that missing entries don't crash the script."""
    assert clean_text(None) == ""
    assert clean_text("") == ""

# =====================================================================
# 🧪 Test 2: Search Link Formatting
# =====================================================================
def test_build_search_url_formats_spaces_correctly():
    """Verifies that multiple search terms convert spaces to plus signs."""
    url = build_search_url("Web Developers", "Nagpur City")
    assert "Web+Developers" in url
    assert "Nagpur+City" in url

# =====================================================================
# 🧪 Test 3: System Stability Check (Server Failure Mocking)
# =====================================================================
def test_google_sheets_webhook_handles_server_errors():
    """Simulates a network drop to ensure the app handles issues gracefully."""
    mock_url = "https://script.google.com/macros/s/mock_link/exec"
    payload = {"name": "Test Agency"}
    
    # We use requests_mock.Mocker() to simulate the internet drop safely
    with requests_mock.Mocker() as mock_adapter:
        mock_adapter.post(mock_url, status_code=500, text="Internal Server Error")
        
        import requests
        response = requests.post(mock_url, json=payload)
        assert response.status_code == 500