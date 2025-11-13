from flask import Flask, render_template, jsonify, request
import imaplib
import email
from email.header import decode_header
from datetime import datetime
import ssl

app = Flask(__name__)
app.config['VERSION'] = '1.0.6'

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


def fetch_emails(imap_server, port, username, password, use_ssl=True, use_tls=False, limit=50):
    """Fetch emails from IMAP server"""
    emails = []
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
                try:
                    date_tuple = email.utils.parsedate_tz(date_str)
                    if date_tuple:
                        date_obj = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                        date_formatted = date_obj.strftime('%Y-%m-%d %H:%M:%S')
                    else:
                        date_formatted = date_str
                except:
                    date_formatted = date_str
                
                # Get email body
                body = ''
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        
                        if content_type == "text/plain" and "attachment" not in content_disposition:
                            try:
                                body = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                                break
                            except:
                                pass
                else:
                    try:
                        body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
                    except:
                        body = str(msg.get_payload())
                
                emails.append({
                    'id': email_id.decode(),
                    'subject': subject,
                    'from': from_addr,
                    'to': to_addr,
                    'date': date_formatted,
                    'body': body[:500] + '...' if len(body) > 500 else body
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


@app.route('/api/version', methods=['GET'])
def get_version():
    """Get application version"""
    return jsonify({'version': app.config['VERSION']})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

