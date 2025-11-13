from flask import Flask, render_template, jsonify, request
import imaplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from datetime import datetime
import ssl
import smtplib
import re

app = Flask(__name__)
app.config['VERSION'] = '1.0.10'

# Default email configurations (can be overridden via settings)
GMAIL_CONFIG = {
    'imap_server': 'imap.gmail.com',
    'port': 993,
    'use_ssl': True,
    'use_tls': False,
    'username': 'eric.brilliant@gmail.com',
    'password': 'opqx pfna kagb bznr'
}

EMAIL163_CONFIG = {
    'imap_server': 'imap.163.com',
    'port': 993,
    'use_ssl': True,
    'use_tls': False,
    'username': '19902475292@163.com',
    'password': 'JDy8MigeNmsESZRa'
}

QQ_CONFIG = {
    'imap_server': 'imap.qq.com',
    'port': 993,
    'use_ssl': True,
    'use_tls': False,
    'username': '',
    'password': ''
}

SMTP_PRIMARY_CONFIG = {
    'name': 'Gmail SMTP',
    'server': 'smtp.gmail.com',
    'port': 587,
    'use_ssl': False,
    'use_tls': True,
    'username': 'eric.brilliant@gmail.com',
    'password': 'opqx pfna kagb bznr',
    'sender_name': 'Mail Task',
    'from_address': 'eric.brilliant@gmail.com'
}

SMTP_BACKUP_CONFIG = {
    'name': '163.com SMTP',
    'server': 'smtp.163.com',
    'port': 465,
    'use_ssl': True,
    'use_tls': False,
    'username': '19902475292@163.com',
    'password': 'JDy8MigeNmsESZRa',
    'sender_name': 'Mail Task Backup',
    'from_address': '19902475292@163.com'
}

DEFAULT_SMTP_CONFIGS = [SMTP_PRIMARY_CONFIG, SMTP_BACKUP_CONFIG]


def decode_mime_words(s):
    """Decode MIME encoded words in email headers"""
    if s is None:
        return ''
    decoded_fragments = decode_header(s)
    decoded_str = ''
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            decoded_str += fragment.decode(encoding or 'utf-8', errors='ignore')
        else:
            decoded_str += fragment
    return decoded_str


def strip_html_tags(text):
    """Remove HTML tags for preview text"""
    if not text:
        return ''
    return re.sub(r'<[^>]+>', '', text)


def build_smtp_config_list(configs):
    """Construct a sanitized list of SMTP configuration dictionaries."""
    sanitized = []
    for cfg in configs:
        if not cfg:
            continue
        server = cfg.get('server')
        username = cfg.get('username')
        password = cfg.get('password')
        if not server or not username or not password:
            continue

        sanitized.append({
            'name': cfg.get('name', server),
            'server': server,
            'port': int(cfg.get('port') or (465 if cfg.get('use_ssl') else 587)),
            'use_ssl': bool(cfg.get('use_ssl')),
            'use_tls': bool(cfg.get('use_tls')),
            'username': username,
            'password': password,
            'timeout': int(cfg.get('timeout') or 10),
            'sender_name': cfg.get('sender_name'),
            'from_address': cfg.get('from_address') or username
        })
    return sanitized


def send_email_with_configs(configs, subject, body, recipients, is_html=False, sender_name=None):
    """Attempt to send email using provided SMTP configs with automatic fallback."""
    attempts = []
    for cfg in configs:
        smtp = None
        try:
            if cfg.get('use_ssl'):
                smtp = smtplib.SMTP_SSL(cfg['server'], cfg['port'], timeout=cfg.get('timeout', 10))
            else:
                smtp = smtplib.SMTP(cfg['server'], cfg['port'], timeout=cfg.get('timeout', 10))
                if cfg.get('use_tls'):
                    smtp.starttls()

            smtp.login(cfg['username'], cfg['password'])

            msg = MIMEText(body or '', 'html' if is_html else 'plain', 'utf-8')
            from_address = cfg.get('from_address') or cfg['username']
            display_name = sender_name or cfg.get('sender_name') or from_address
            msg['From'] = email.utils.formataddr((display_name, from_address))
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = subject or ''

            smtp.sendmail(from_address, recipients, msg.as_string())
            smtp.quit()
            return {'success': True, 'provider': cfg.get('name', cfg['server'])}
        except Exception as exc:
            attempts.append({'provider': cfg.get('name', cfg['server']), 'error': str(exc)})
            if smtp:
                try:
                    smtp.quit()
                except Exception:
                    pass
    return {'success': False, 'errors': attempts}



def fetch_emails(imap_server, port, username, password, use_ssl=True, use_tls=False, limit=50):
    """Fetch emails from IMAP server"""
    emails = []
    today = datetime.now().date()
    try:
        # Create SSL context
        context = ssl.create_default_context()
        
        # Connect to IMAP server based on SSL/TLS settings
        if use_ssl:
            mail = imaplib.IMAP4_SSL(imap_server, port, ssl_context=context)
        else:
            mail = imaplib.IMAP4(imap_server, port)
            if use_tls:
                mail.starttls()
        
        mail.login(username, password)
        mail.select('INBOX')
        
        # Search for all emails
        status, messages = mail.search(None, 'ALL')
        email_ids = messages[0].split()
        
        # Get the most recent emails (limit)
        email_ids = email_ids[-limit:] if len(email_ids) > limit else email_ids
        
        for email_id in reversed(email_ids):
            try:
                status, msg_data = mail.fetch(email_id, '(RFC822)')
                if status != 'OK':
                    continue
                
                email_body = msg_data[0][1]
                msg = email.message_from_bytes(email_body)
                
                # Extract email details
                subject = decode_mime_words(msg.get('Subject', ''))
                from_addr = decode_mime_words(msg.get('From', ''))
                to_addr = decode_mime_words(msg.get('To', ''))
                date_str = msg.get('Date', '')
                
                # Parse date
                date_obj = None
                date_formatted = date_str
                try:
                    parsed_dt = email.utils.parsedate_to_datetime(date_str)
                    if parsed_dt:
                        if parsed_dt.tzinfo is not None:
                            parsed_dt = parsed_dt.astimezone().replace(tzinfo=None)
                        date_obj = parsed_dt
                        date_formatted = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                except Exception:
                    pass

                if not date_obj or date_obj.date() != today:
                    continue

                # Get email body (plain and html)
                body_plain = ''
                body_html = ''

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))

                        if content_type == "text/plain" and "attachment" not in content_disposition and not body_plain:
                            try:
                                body_plain = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            except Exception:
                                body_plain = ''
                        elif content_type == "text/html" and "attachment" not in content_disposition and not body_html:
                            try:
                                body_html = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                            except Exception:
                                body_html = ''
                else:
                    content_type = msg.get_content_type()
                    try:
                        payload = msg.get_payload(decode=True)
                        decoded = payload.decode('utf-8', errors='ignore') if isinstance(payload, bytes) else str(payload)
                    except Exception:
                        decoded = str(msg.get_payload())

                    if content_type == "text/html":
                        body_html = decoded
                    else:
                        body_plain = decoded

                preview_source = (body_plain or strip_html_tags(body_html) or '').strip()
                preview = preview_source[:500] + '...' if len(preview_source) > 500 else preview_source

                emails.append({
                    'id': email_id.decode(),
                    'subject': subject,
                    'from': from_addr,
                    'to': to_addr,
                    'date': date_formatted,
                    'preview': preview,
                    'plain_body': body_plain,
                    'html_body': body_html
                })
            except Exception as e:
                print(f"Error processing email {email_id}: {str(e)}")
                continue
        
        mail.close()
        mail.logout()
        
    except Exception as e:
        print(f"Error fetching emails: {str(e)}")
        return {'error': str(e)}
    
    return {'emails': emails, 'count': len(emails)}


@app.route('/')
def index():
    return render_template('index.html', version=app.config['VERSION'])


@app.route('/api/fetch-gmail', methods=['POST'])
def fetch_gmail():
    """Fetch emails from Gmail"""
    data = request.json or {}
    limit = data.get('limit', 50)
    
    # Use provided config or default
    config = {
        'imap_server': data.get('imap_server', GMAIL_CONFIG['imap_server']),
        'port': data.get('port', GMAIL_CONFIG['port']),
        'username': data.get('username', GMAIL_CONFIG['username']),
        'password': data.get('password', GMAIL_CONFIG['password']),
        'use_ssl': data.get('use_ssl', GMAIL_CONFIG.get('use_ssl', True)),
        'use_tls': data.get('use_tls', GMAIL_CONFIG.get('use_tls', False))
    }
    
    result = fetch_emails(
        config['imap_server'],
        config['port'],
        config['username'],
        config['password'],
        config['use_ssl'],
        config['use_tls'],
        limit
    )
    return jsonify(result)


@app.route('/api/fetch-qq', methods=['POST'])
def fetch_qq():
    """Fetch emails from QQ"""
    data = request.json
    username = data.get('username', '')
    password = data.get('password', '')
    
    if not username or not password:
        return jsonify({'error': 'QQ username and password are required'}), 400
    
    limit = data.get('limit', 50)
    result = fetch_emails(
        data.get('imap_server', QQ_CONFIG['imap_server']),
        data.get('port', QQ_CONFIG['port']),
        username,
        password,
        data.get('use_ssl', QQ_CONFIG.get('use_ssl', True)),
        data.get('use_tls', QQ_CONFIG.get('use_tls', False)),
        limit
    )
    return jsonify(result)


@app.route('/api/fetch-163', methods=['POST'])
def fetch_163():
    """Fetch emails from 163.com"""
    data = request.json or {}
    limit = data.get('limit', 50)
    
    # Use provided config or default
    config = {
        'imap_server': data.get('imap_server', EMAIL163_CONFIG['imap_server']),
        'port': data.get('port', EMAIL163_CONFIG['port']),
        'username': data.get('username', EMAIL163_CONFIG['username']),
        'password': data.get('password', EMAIL163_CONFIG['password']),
        'use_ssl': data.get('use_ssl', EMAIL163_CONFIG.get('use_ssl', True)),
        'use_tls': data.get('use_tls', EMAIL163_CONFIG.get('use_tls', False))
    }
    
    result = fetch_emails(
        config['imap_server'],
        config['port'],
        config['username'],
        config['password'],
        config['use_ssl'],
        config['use_tls'],
        limit
    )
    return jsonify(result)


@app.route('/api/send-email', methods=['POST'])
def send_email():
    """Send email using configured SMTP servers with automatic fallback."""
    data = request.json or {}
    recipients_raw = data.get('to') or data.get('recipients')

    if not recipients_raw:
        return jsonify({'error': 'Recipient email address is required'}), 400

    if isinstance(recipients_raw, list):
        recipients = [addr.strip() for addr in recipients_raw if isinstance(addr, str) and addr.strip()]
    else:
        recipients = [addr.strip() for addr in re.split(r'[;,]', str(recipients_raw)) if addr.strip()]

    if not recipients:
        return jsonify({'error': 'Recipient email address is required'}), 400

    configs_payload = data.get('configs') or []
    configs = build_smtp_config_list(configs_payload)
    if not configs:
        configs = build_smtp_config_list(DEFAULT_SMTP_CONFIGS)

    result = send_email_with_configs(
        configs,
        data.get('subject', ''),
        data.get('body', ''),
        recipients,
        bool(data.get('is_html')),
        data.get('sender_name')
    )

    if result.get('success'):
        return jsonify({'status': 'sent', 'provider': result.get('provider')})

    return jsonify({
        'error': 'Unable to send email via configured SMTP servers',
        'details': result.get('errors', [])
    }), 500


@app.route('/api/version', methods=['GET'])
def get_version():
    """Get application version"""
    return jsonify({'version': app.config['VERSION']})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

