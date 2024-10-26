import pytest
import requests
from teams import get_access_token, get_outlook_emails, send_outlook_email, get_teams_messages

# Mocking the environment variables
import os
from unittest.mock import patch

@pytest.fixture
def mock_env_vars():
    # Mock environment variables with realistic fake data
    with patch.dict(os.environ, {
        'TENANT_ID': 'fake-tenant-id',
        'CLIENT_ID': 'fake-client-id',
        'CLIENT_SECRET': 'fake-client-secret',
        'AUTHORITY': 'https://login.microsoftonline.com/fake-tenant-id',
        'USER_EMAIL': 'test@example.com',
    }):
        yield


@pytest.fixture
def mock_requests(requests_mock):
    # This fixture uses requests_mock to mock HTTP requests
    return requests_mock


# Test for get_access_token()
def test_get_access_token(mock_env_vars):
    with patch("msal.ConfidentialClientApplication.acquire_token_for_client") as mock_acquire_token:
        # Mock the token response
        mock_acquire_token.return_value = {"access_token": "fake_token"}

        access_token = get_access_token()
        assert access_token == "fake_token"


# Test for get_outlook_emails()
def test_get_outlook_emails(mock_env_vars, mock_requests):
    # Mock the GET request to fetch Outlook emails
    mock_requests.get("https://graph.microsoft.com/v1.0/users/test@example.com/messages", json={
        "value": [{"subject": "Test email", "from": {"emailAddress": {"address": "sender@example.com"}}}]
    })
    
    # Call the function and assert the results
    emails = get_outlook_emails("fake_token", "test@example.com")
    
    assert len(emails) == 1
    assert emails[0]['subject'] == "Test email"
    assert emails[0]['from']['emailAddress']['address'] == "sender@example.com"


# Test for send_outlook_email()
def test_send_outlook_email(mock_env_vars, mock_requests):
 with patch("teams.USER_EMAIL", "test@example.com"): 
    # Mock the POST request to send an email
    mock_requests.post("https://graph.microsoft.com/v1.0/users/test@example.com/sendMail", status_code=202)
    
    # Call the function
    send_outlook_email("fake_token", "Test Subject", "Test Body", ["recipient@example.com"])
    
    # Assert that the mock POST request was called
    assert mock_requests.called


# Test for get_teams_messages()
def test_get_teams_messages(mock_env_vars, mock_requests):
 with patch("teams.USER_EMAIL", "test@example.com"):
    # Mock the GET request to fetch Teams messages
    mock_requests.get("https://graph.microsoft.com/v1.0/users/test@example.com/chats", json={
        "value": [{"id": "123", "createdDateTime": "2024-10-26T10:00:00Z"}]
    })
    
    # Call the function and assert the results
    messages = get_teams_messages("fake_token")
    
    assert len(messages) == 1
    assert messages[0]['id'] == "123"
    assert messages[0]['createdDateTime'] == "2024-10-26T10:00:00Z"
