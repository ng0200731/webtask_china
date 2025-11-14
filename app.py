from flask import Flask, render_template, jsonify, request
import imaplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from datetime import datetime, date
from typing import Optional
import ssl
import smtplib
import re
import os
from pathlib import Path
import sqlite3

app = Flask(__name__)
app.config['VERSION'] = '1.0.27'

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
CUSTOMER_DB_PATH = Path(__file__).resolve().parent / 'mailtask.db'


def get_db_connection():
    connection = sqlite3.connect(str(CUSTOMER_DB_PATH))
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    connection = None
    cursor = None
    try:
        connection = get_db_connection()
        cursor = connection.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                email_suffix TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                provider TEXT NOT NULL,
                email_uid TEXT NOT NULL,
                subject TEXT,
                from_addr TEXT,
                to_addr TEXT,
                date TEXT,
                preview TEXT,
                plain_body TEXT,
                html_body TEXT,
                sequence TEXT,
                fetched_at TEXT DEFAULT (datetime('now')),
                UNIQUE(provider, email_uid)
            )
        """)
        connection.commit()
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()


def insert_customer(name: str, email_suffix: str) -> int:
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        "INSERT INTO customers (name, email_suffix) VALUES (?, ?)",
        (name, email_suffix)
    )
    connection.commit()
    customer_id = cursor.lastrowid
    cursor.close()
    connection.close()
    return customer_id


def fetch_customers():
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("""
        SELECT id, name, email_suffix, created_at
        FROM customers
        ORDER BY datetime(created_at) DESC
    """)
    rows = cursor.fetchall()
    cursor.close()
    connection.close()

    customers = []
    for row in rows:
        created_at = row['created_at']
        if created_at:
            try:
                created_at_iso = datetime.fromisoformat(created_at.replace(' ', 'T'))
                created_at = created_at_iso.isoformat()
            except ValueError:
                pass

        customers.append({
            'id': row['id'],
            'name': row['name'],
            'email_suffix': row['email_suffix'],
            'created_at': created_at
        })
    return customers


def save_emails(provider: str, emails: list[dict]):
    if not emails:
        return
    connection = get_db_connection()
    cursor = connection.cursor()
    now_iso = datetime.utcnow().isoformat()
    rows = []
    for email in emails:
        email_uid = str(email.get('id') or email.get('email_uid') or '')
        if not email_uid:
            continue
        rows.append((
            provider,
            email_uid,
            email.get('subject'),
            email.get('from'),
            email.get('to'),
            email.get('date'),
            email.get('preview'),
            email.get('plain_body'),
            email.get('html_body'),
            email.get('sequence'),
            now_iso
        ))
    if not rows:
        cursor.close()
        connection.close()
        return

    cursor.executemany("""
        INSERT INTO emails (
            provider,
            email_uid,
            subject,
            from_addr,
            to_addr,
            date,
            preview,
            plain_body,
            html_body,
            sequence,
            fetched_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(provider, email_uid) DO UPDATE SET
            subject = excluded.subject,
            from_addr = excluded.from_addr,
            to_addr = excluded.to_addr,
            date = excluded.date,
            preview = excluded.preview,
            plain_body = excluded.plain_body,
            html_body = excluded.html_body,
            sequence = excluded.sequence,
            fetched_at = excluded.fetched_at
    """, rows)
    connection.commit()
    cursor.close()
    connection.close()


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


def build_sequence_code(from_address: str, email_date: Optional[datetime] = None) -> str:
    """Construct a sequence code YYYYMMDD_<two letters before @>_<domain label>."""
    sequence_date = (email_date or datetime.now()).strftime('%Y%m%d')
    parsed_email = email.utils.parseaddr(from_address)[1].lower() if from_address else ''
    local_part = ''
    domain_part = ''

    if parsed_email and '@' in parsed_email:
        local_part, domain_part = parsed_email.split('@', 1)
    elif parsed_email:
        local_part = parsed_email

    letters = [ch for ch in local_part if ch.isalpha()]
    if len(letters) >= 2:
        prefix = ''.join(letters[:2]).lower()
    elif len(letters) == 1:
        prefix = letters[0].lower() + 'x'
    else:
        prefix = 'xx'

    domain_label = ''
    if domain_part:
        first_label = domain_part.split('.')[0]
        domain_label = ''.join(ch for ch in first_label if ch.isalnum()).lower()
    if not domain_label:
        domain_label = 'domain'

    return f'{sequence_date}_{prefix}_{domain_label}'


def fetch_emails(imap_server, port, username, password, use_ssl=True, use_tls=False, limit=50):
    """Fetch emails from IMAP server"""
    emails = []
    today = datetime.now().date()
    def _imap_date(d: date) -> str:
        return d.strftime('%d-%b-%Y')
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
        
        # Search for today's emails on the server side to reduce bandwidth
        # Use SINCE today and (optionally) BEFORE tomorrow for stricter bounds
        since_clause = f'(SINCE { _imap_date(today) })'
        status, messages = mail.search(None, since_clause)
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
                sequence_code = build_sequence_code(from_addr, date_obj)

                emails.append({
                    'id': email_id.decode(),
                    'subject': subject,
                    'from': from_addr,
                    'to': to_addr,
                    'date': date_formatted,
                    'preview': preview,
                    'plain_body': body_plain,
                    'html_body': body_html,
                    'sequence': sequence_code
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
    try:
        save_emails('gmail', result.get('emails', []))
    except Exception as exc:
        print(f"Error saving Gmail emails: {exc}")
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
    try:
        save_emails('qq', result.get('emails', []))
    except Exception as exc:
        print(f"Error saving QQ emails: {exc}")
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
    try:
        save_emails('163', result.get('emails', []))
    except Exception as exc:
        print(f"Error saving 163 emails: {exc}")
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


@app.route('/api/customers', methods=['GET', 'POST'])
def customers_endpoint():
    if request.method == 'GET':
        try:
            customers = fetch_customers()
            return jsonify({'customers': customers})
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500

    data = request.json or {}
    name = (data.get('name') or '').strip()
    suffix = (data.get('email_suffix') or '').strip()

    if not name:
        return jsonify({'error': 'Customer name is required'}), 400
    if not suffix:
        return jsonify({'error': 'Email suffix is required'}), 400
    if '@' in suffix:
        return jsonify({'error': 'Do not include "@" in the email suffix'}), 400
    if not re.match(r'^[a-zA-Z0-9._-]+\.[a-zA-Z]{2,}$', suffix):
        return jsonify({'error': 'Invalid email suffix format'}), 400

    normalized_suffix = '@' + suffix

    try:
        customer_id = insert_customer(name, normalized_suffix)
        return jsonify({'id': customer_id, 'name': name, 'email_suffix': normalized_suffix}), 201
    except Exception as exc:
        return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/emails', methods=['GET', 'POST'])
def handle_emails():
    if request.method == 'GET':
        provider = (request.args.get('provider') or '').strip().lower()
        if not provider:
            return jsonify({'error': 'Provider query parameter is required'}), 400

        today_iso = datetime.now().date().isoformat()
        connection = get_db_connection()
        cursor = connection.cursor()
        cursor.execute("""
            SELECT provider, email_uid, subject, from_addr, to_addr, date, preview, plain_body, html_body, sequence, fetched_at
            FROM emails
            WHERE provider = ?
              AND date(datetime(fetched_at)) = ?
            ORDER BY datetime(date) DESC, datetime(fetched_at) DESC
        """, (provider, today_iso))
        rows = cursor.fetchall()
        cursor.close()
        connection.close()

        emails = []
        for row in rows:
            emails.append({
                'id': row[1],
                'subject': row[2],
                'from': row[3],
                'to': row[4],
                'date': row[5],
                'preview': row[6],
                'plain_body': row[7],
                'html_body': row[8],
                'sequence': row[9],
                'fetched_at': row[10]
            })

        return jsonify({'provider': provider, 'emails': emails})
    
    elif request.method == 'POST':
        data = request.json or {}
        provider = (data.get('provider') or '').strip().lower()
        emails = data.get('emails') or []
        target_date = data.get('date')  # Optional: YYYY-MM-DD format
        
        if not provider:
            return jsonify({'error': 'Provider is required'}), 400
        
        if not isinstance(emails, list):
            return jsonify({'error': 'Emails must be a list'}), 400
        
        # Save emails (save_emails handles the date from email.date field)
        save_emails(provider, emails)
        return jsonify({'status': 'saved', 'count': len(emails)})


@app.route('/api/version', methods=['GET'])
def get_version():
    """Get application version"""
    return jsonify({'version': app.config['VERSION']})


initialize_database()


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

