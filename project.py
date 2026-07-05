import os
import time
import requests
import openai
import random
from dotenv import load_dotenv
import win32com.client
import pywintypes
from openai import OpenAI


def main():
    """Orchestrate the full continuous email automation workflow using Outlook COM."""
    load_dotenv()
    config = load_config()
    outlook = connect_outlook()
 

    try:
        while True:
            focus_outlook(outlook)
            messages = fetch_unread_emails(outlook)
            print("Current Unread Issues as of " + time.strftime("%Y-%m-%d %H:%M") + ": " + str(len(messages)) + "\n")

            for message in messages:
                if is_reply(message) and "The Following Ticket has been Created:" in message.Body:
                    print("Ignoring reply email with ticket creation confirmation: \n")
                    mark_email_as_processed(outlook,message)
                    print("Ticket already processed and moved to ""AI Response/Handled Already"" folder.\n")

                if is_reply(message) and not "The Following Ticket has been Created:" in message.Body:
                    print("Processing Reply Email " + str(messages.index(message) + 1) + " of " + str(len(messages)) + "\n")
                    print("Subject: " + message.Subject + "\n" + "Body: " + message.Body + "\n" + "Sender Email: " + message.SenderEmailAddress + "\n" + "Sender Name: " + message.SenderName + "\n" + "Unread: " + str(message.UnRead) + "\n")
                    time.sleep(config.get("poll_interval", 15))
                    generated_ticket_num = random.randint(1000, 9999)
                    ticket_reply_body = ticket_response(message.SenderName, generated_ticket_num)
                    print("Reply:" + ticket_reply_body + "\n")
                    send_email(
                        outlook,
                        message.SenderEmailAddress,
                        f"Re: Ticket #{generated_ticket_num} Created - {message.Subject}",
                        ticket_reply_body,
                    )
                    print("Email sent to " + message.SenderEmailAddress + "\n")
                    ticket_processed(outlook,message)
                    print("Ticket processed and moved to ""Open Cases"" folder.\n")

                if "Issue:" in message.Subject and not is_reply(message):
                    print("Processing New Email " + str(messages.index(message) + 1) + " of " + str(len(messages)) + "\n")
                    print("Subject: " + message.Subject + "\n" + "Body: " + message.Body + "\n" + "Sender Email: " + message.SenderEmailAddress + "\n" + "Sender Name: " + message.SenderName + "\n" + "Unread: " + str(message.UnRead) + "\n")
                    print("prompt:" + message.Body + "\n")
                    response_text = query_chatgpt(message.Body, config["openai_api_key"])
                    print("ChatGPT Response: " + response_text + "\n")
                    time.sleep(config.get("poll_interval", 15))
                    reply_body = draft_response(message.SenderName, response_text)
                    print("Reply:" + reply_body + "\n")
                    send_email(
                        outlook,
                        message.SenderEmailAddress,
                        f"Re: {message.Subject}",
                        reply_body,
                    )
                    print("Email sent to " + message.SenderEmailAddress + "\n")
                    mark_email_as_processed(outlook,message)
                    print("Email marked as read and moved to ""AI Response/Handled Already"" folder.\n")
            print("--------------------------------------------------\n")

            time.sleep(config.get("poll_interval", 15))
    except KeyboardInterrupt:
        print("Automation stopped.")


def load_config():
    """Load environment-based configuration."""
    return {
        "openai_api_key": os.getenv("OPENAI_API_KEY"),
        "poll_interval": int(os.getenv("POLL_INTERVAL", "60")),
    }


def connect_outlook():
    """Connect to Outlook via COM automation."""
    # Get the Outlook application object.
    # Returns the Outlook.Application COM object.
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook


def focus_outlook(outlook):
    """Bring Outlook to the foreground before refreshing the Inbox."""
    try:
        outlook.ActiveExplorer().Activate()
    except Exception as exc:
        print(f"Outlook focus warning: {exc}")


def fetch_unread_emails(outlook):
    """Fetch unread emails from the Inbox and force a refresh first."""
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

    focus_outlook(outlook)
    time.sleep(1)

    try:
        inbox.Refresh()
    except Exception as exc:
        print(f"Inbox refresh warning: {exc}")

    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    unread_messages = []
    for message in items:
        if message.UnRead:
            unread_messages.append(message)

    return unread_messages



def query_chatgpt(issue_text, api_keys):
    """Send the prompt to ChatGPT / OpenAI API and return the response."""
    # Use requests or the openai client to make an API call.
    # Return the model's generated text.
    client = OpenAI(api_key=api_keys)

    response = client.responses.create(
        model="gpt-5-mini",
        instructions="""
    You are a helpdesk support agent. your role is to do the following based on the following email message recieved from the user:
    1. Ignore all emotional input from the email sender.
    2. Emphasize how frustrating the experience can be.
    3. only Provide the steps to resolve the issue
    4. Do not Greet them
    5. Do not ask them to reply with any additional information
    6. Do not ask them to confirm if the issue is resolved
    7. Keep responses concise.
    """,
        input=issue_text
    )

    return response.output_text


def draft_response(original_sender, chatgpt_response):
    """Create a polite email reply using the ChatGPT answer."""
    return (
        f"Hello {original_sender},\n\n"
        f"Thank you for reaching out. Based on your issue, here is a suggested solution:\n\n"
        f"{chatgpt_response}\n\n"
        f"If the Issue still presists, please reply to this email for an automatic ticket creation with ONLY the following details:\n\n"
        f"User name: \n"
        f"Device type (desktop, mobile): \n"
        f"Device Name: \n"
        f"Application name: \n"
        f"Version: \n"
        f"Browser (Chrome, Edge, Firefox): \n"
        f"Operating System: \n\n"
        f"Best regards,\n\n"
        f"CS50 IT Support Automation"
    )

def ticket_response(original_sender, ticket_number):
    return (
        f"Hello {original_sender},\n\n"
        f"Thank you for your reply with the Device/Application details causing the issue.\n\n"
        f"The Following Ticket has been Created: #{ticket_number}\n\n"
        f"Kindly keep this ticket number for your records and future reference.\n\n"
        f"Best regards,\n\n"
        f"CS50 IT Support Automation"
    )

def send_email(outlook, recipient, subject, body):
    """Send an email reply using Outlook."""
    # Create a new mail item.
    # Set recipient, subject, and body.
    # Send the message.
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = recipient
    mail.Subject = subject
    mail.Body = body
    mail.Send()

def is_reply(message):
    subject = (message.Subject or "").strip().lower()
    if subject.startswith("re:"):
        return True
    if subject.startswith("fw:") or subject.startswith("fwd:"):
        return True
    return False

def mark_email_as_processed(outlook,message):
    """Mark an email as read after processing."""
    # Set the UnRead flag to False.
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.GetDefaultFolder(6).Folders["AI Response/Handled Already"]
    message.UnRead = False
    message.Move(folder)

def ticket_processed(outlook,message):
    """Mark an email as read after processing."""
    # Set the UnRead flag to False.
    namespace = outlook.GetNamespace("MAPI")
    folder = namespace.GetDefaultFolder(6).Folders["Open Cases"]
    message.UnRead = False
    message.Move(folder)

if __name__ == "__main__":
    main()