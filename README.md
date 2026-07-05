# AI Powered Enterprise Helpdesk Ticket Handling/Case Creation Automation

#### Video Demo: https://youtu.be/9cnDvsXFvG0
#### Description:
This project is a proof-of-concept help desk automation system that uses AI to process incoming support emails, draft responses, and support a ticket-based workflow. The current version demonstrates the core idea locally on Windows using Microsoft Outlook Classic, while the long-term goal is to evolve that prototype into a scalable enterprise solution.

## What the project is
This project automates a simplified IT support workflow. When a new support email arrives, the system reads the message, sends the content to an AI model, and generates a helpful response. If the sender follows up with more troubleshooting details, the system recognizes that as a reply and handles it as a ticket-related interaction instead of treating it as a brand-new request. The goal is to make support operations faster, more consistent, and more scalable.

## What each file contains and does
- **project.py**: Contains the main automation logic. It loads configuration values, connects to Outlook, scans the inbox for unread messages, classifies messages, generates AI responses, sends replies, and moves processed emails into Outlook folders.
- **test_project.py**: Contains unit tests for the key helper functions. These tests verify that environment-based configuration works, that reply-style subjects are recognized correctly, and that response text includes the expected support messaging.
- **requirements.txt**: Lists the Python dependencies required to run the project, including packages for Outlook automation, HTTP requests, OpenAI access, and environment variables.
- **README.md**: Documents the project, explains its purpose, and outlines both the current implementation and the future enterprise architecture.

## Current state solution
The current solution is a local Windows-based prototype. It uses Microsoft Outlook Classic through the COM automation interface so that the script can interact with the desktop Outlook client directly. The workflow is intentionally simple:

1. The application checks the inbox for unread messages.
2. It identifies whether the email is a new request or a follow-up reply.
3. It sends the message content to an OpenAI model.
4. It drafts a response and sends it from Outlook.
5. It moves processed messages into folders such as "AI Response/Handled Already" or "Open Cases".

This version is useful for demonstration and testing, but it depends on a Windows desktop environment and a polling loop that checks the inbox at intervals.

## Ideal enterprise future state solution
The ideal future version would be a cloud-native, event-driven support platform. Instead of polling the inbox repeatedly, the application would respond to mailbox events as they happen. It would use Microsoft Graph API rather than Outlook Classic COM so that it can work with Microsoft 365 mail in a modern and scalable way.

In that future architecture:
- messages would trigger processing immediately through event subscriptions or webhooks,
- the system would run in containers using Docker for consistent deployment,
- it could be hosted on Azure, AWS, or Google Cloud Platform,
- authentication would use Microsoft Entra ID and secure service principals or managed identities,
- tickets would be created directly in tools such as Jira or ServiceNow instead of only generating an email acknowledgement,
- the AI layer would support both customer communication and backend workflow automation.

## Design choices and why
Several design choices were made to balance practicality with long-term vision.

- **Why use OpenAI?**: The AI layer provides high-quality, consistent support language without requiring a fully custom NLP model.
- **Why use Outlook COM in the current version?**: It offered a workable way to interact with Outlook on Windows for a prototype when easier integrations were not available.
- **Why use polling now?**: Polling is simple to implement and is appropriate for a basic proof of concept.
- **Why move to event-driven architecture later?**: Event-driven processing is more responsive, efficient, and appropriate for real-world support systems.
- **Why use Docker and cloud hosting later?**: Containerization and cloud deployment make the system easier to operate consistently across different environments.
- **Why replace Outlook Classic with Microsoft Graph?**: Microsoft Graph is the modern, supported path for Microsoft 365 integrations and removes the dependency on desktop Outlook behavior.
- **Why connect to Jira or ServiceNow?**: This creates a true enterprise workflow by turning support requests into tracked operational tickets rather than leaving the work inside email.

Overall, the current version is a working prototype, while the future version is designed to become a production-ready intelligent support platform that is scalable, secure, and cloud-friendly.