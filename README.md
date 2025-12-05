# Email Receiver Web App

A web application for receiving emails from Gmail and QQ email accounts.

## Version
v1.0.70

## Gmail API Setup (Required for Attachments)

To fetch Gmail emails with attachments, you need to set up OAuth 2.0:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select existing one
3. Enable Gmail API
4. Go to "Credentials" → "Create Credentials" → "OAuth 2.0 Client ID"
5. Choose "Web Application"
6. Add authorized redirect URI: `http://localhost:5000/oauth2callback`
7. Copy Client ID and Client Secret
8. In the app Settings, enter these credentials in the "Gmail API OAuth 2.0" section
9. Click "Authenticate Gmail" to complete OAuth flow
10. Now you can fetch Gmail emails with attachments!

## Features
- Receive Gmail and QQ emails for both today and yesterday
- Simple black and white UI with Arial font
- Left menu collapses to initials (~10% width) and expands on hover
- Right content area displays today's emails with rich HTML preview
- Click an email to open the original HTML content in a modal viewer
- Built-in SMTP sender with Gmail primary config and 163.com automatic fallback
- Caches today's fetched emails locally and restores them on page load
- Persists today's and yesterday's emails (including attachments) to SQLite
- Generates per-email sequence codes (`YYYYMMDD_xx_domain`) for quick identification
- Manage customer directory stored in a local SQLite database (name + email suffix, no external server required)

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Run the application:
```bash
python app.py
```

3. Open your browser and navigate to:
```
http://localhost:5000
```

## Usage

1. Click "Receive Mail" in the left menu
2. For Gmail: Click "Fetch Gmail" button (credentials are pre-configured)
3. For QQ Email: 
   - Enter your QQ email address and authorization code
   - Click "Save Credentials"
   - Click "Fetch QQ Email" button
4. Click "Send" in the left menu, fill in recipient(s), subject, body, and click "Send Email"

## Configuration

Gmail credentials are pre-configured in `app.py`. For QQ email, you need to:
1. Enable IMAP in your QQ email settings
2. Generate an authorization code
3. Enter these credentials in the web interface
