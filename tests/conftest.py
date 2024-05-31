import os

import pytest

FILE_PATH = os.path.dirname(os.path.abspath(__file__))


@pytest.fixture()
def msg():
    return "Test message"


@pytest.fixture()
def sender_email():
    return "test-sender@example.com"


@pytest.fixture()
def subject():
    return "Test subject"


@pytest.fixture()
def receiver_emails():
    return [
        "test-receiver@example.com",
    ]


@pytest.fixture()
def attachments():
    return [
        os.path.join(FILE_PATH, "test_data", "test_documents", d)
        for d in os.listdir(os.path.join(FILE_PATH, "test_data", "test_documents"))
    ]
