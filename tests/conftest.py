import pytest


@pytest.fixture
def msg():
    return "Test message"


@pytest.fixture
def sender_email():
    return "test-sender@example.com"


@pytest.fixture
def subject():
    return "Test subject"

    
@pytest.fixture
def receiver_emails():
    return [
        "test-receiver@example.com",
    ]
        
@pytest.fixture
def attachments():
    return []
