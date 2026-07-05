from project import draft_response, is_reply, load_config, ticket_response


def test_load_config(monkeypatch):
    # Tests that load_config reads environment variables and converts the interval to an int.
    monkeypatch.setenv("OPENAI_API_KEY", "test-key")
    monkeypatch.setenv("POLL_INTERVAL", "45")

    config = load_config()

    assert config["openai_api_key"] == "test-key"  # load_config should read OPENAI_API_KEY from the environment
    assert config["poll_interval"] == 45  # load_config should convert POLL_INTERVAL to an integer


def test_is_reply():
    # Tests that is_reply returns True for reply-style subjects and False for normal subjects.
    class Message:
        Subject = "Re: issue"

    assert is_reply(Message()) is True  # is_reply should recognize reply-style subjects like 'Re:'

    class OtherMessage:
        Subject = "New issue"

    assert is_reply(OtherMessage()) is False  # is_reply should return False for non-reply subjects


def test_draft_response_and_ticket_response():
    # Tests that the response helpers include the expected greeting, message content, and ticket details.
    draft = draft_response("Ada", "Step 1")
    assert "Hello Ada" in draft  # draft_response should greet the sender
    assert "CS50 IT Support Automation" in draft  # draft_response should include the support signature

    ticket = ticket_response("Ada", 1234)
    assert "Hello Ada" in ticket  # ticket_response should greet the sender
    assert "The Following Ticket has been Created: #1234" in ticket  # ticket_response should include the generated ticket number