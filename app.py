from flask import Flask, render_template, jsonify, request, redirect, session, url_for
import imaplib
import email
from email.header import decode_header
from email.mime.text import MIMEText
from datetime import datetime, date, timedelta
from typing import Optional
import ssl
import smtplib
import re
import os
import json
import base64
from pathlib import Path
import sqlite3
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

app = Flask(__name__)
app.config['VERSION'] = '1.0.66'
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')

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

# Gmail API OAuth 2.0 Configuration
# Users need to set these in Settings or environment variables
GMAIL_OAUTH_CONFIG = {
    'client_id': os.environ.get('GMAIL_CLIENT_ID', ''),
    'client_secret': os.environ.get('GMAIL_CLIENT_SECRET', ''),
    'redirect_uri': os.environ.get('GMAIL_REDIRECT_URI', 'http://localhost:5000/oauth2callback'),
    'scopes': ['https://www.googleapis.com/auth/gmail.readonly']
}


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
                country TEXT,
                website TEXT,
                remark TEXT,
                attachments TEXT,
                company_name TEXT,
                tel TEXT,
                source TEXT,
                address TEXT,
                business_type TEXT,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Migrate existing customers table to add new columns if they don't exist
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN country TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN website TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN remark TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN attachments TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN company_name TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN tel TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN source TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN address TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE customers ADD COLUMN business_type TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
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
                attachments TEXT,
                fetched_at TEXT DEFAULT (datetime('now')),
                UNIQUE(provider, email_uid)
            )
        """)
        cursor.execute("PRAGMA table_info(emails)")
        email_columns = {row['name'] for row in cursor.fetchall()}
        if 'attachments' not in email_columns:
            cursor.execute("ALTER TABLE emails ADD COLUMN attachments TEXT")
        
        # OAuth tokens table for Gmail API
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS oauth_tokens (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                provider TEXT NOT NULL UNIQUE,
                token TEXT,
                refresh_token TEXT,
                token_uri TEXT,
                client_id TEXT,
                client_secret TEXT,
                scopes TEXT,
                updated_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # OAuth states table for state parameter validation
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS oauth_states (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                state TEXT NOT NULL UNIQUE,
                client_id TEXT,
                client_secret TEXT,
                redirect_uri TEXT,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Clean up old states (older than 10 minutes)
        cursor.execute("""
            DELETE FROM oauth_states 
            WHERE datetime(created_at) < datetime('now', '-10 minutes')
        """)
        # Task types table for dropdown options
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS task_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                display_order INTEGER DEFAULT 0,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Initialize default task types if table is empty
        cursor.execute("SELECT COUNT(*) as count FROM task_types")
        if cursor.fetchone()['count'] == 0:
            default_types = [
                ('need sample', 1),
                ('quotation', 2),
                ('outsource', 3)
            ]
            cursor.executemany("""
                INSERT INTO task_types (name, display_order) VALUES (?, ?)
            """, default_types)
        
        # Tasks table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sequence TEXT,
                customer TEXT,
                email TEXT,
                catalogue TEXT NOT NULL,
                template TEXT NOT NULL,
                attachments TEXT,
                deadline TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now'))
            )
        """)
        try:
            cursor.execute("ALTER TABLE tasks ADD COLUMN deadline TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        try:
            cursor.execute("ALTER TABLE tasks ADD COLUMN email TEXT")
        except sqlite3.OperationalError:
            pass  # Column already exists
        cursor.execute("PRAGMA table_info(tasks)")
        task_columns = {row['name'] for row in cursor.fetchall()}
        if 'updated_at' not in task_columns:
            cursor.execute("ALTER TABLE tasks ADD COLUMN updated_at TEXT")
            cursor.execute("UPDATE tasks SET updated_at = created_at WHERE updated_at IS NULL")
        # Countries table for dropdown options
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS countries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                display_order INTEGER DEFAULT 0,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Customer sources table for dropdown options
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS customer_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                display_order INTEGER DEFAULT 0,
                created_at TEXT DEFAULT (datetime('now'))
            )
        """)
        # Initialize default customer sources if table is empty
        cursor.execute("SELECT COUNT(*) as count FROM customer_sources")
        if cursor.fetchone()['count'] == 0:
            default_sources = [
                ('2025 PT Jakarta Intl Expo', 1),
                ('Sales Referral', 2)
            ]
            cursor.executemany("""
                INSERT INTO customer_sources (name, display_order) VALUES (?, ?)
            """, default_sources)
        
        # Customer business types table for dropdown options
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS customer_business_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            display_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now'))
        )
        """)
        # Initialize default customer business types if table is empty
        cursor.execute("SELECT COUNT(*) as count FROM customer_business_types")
        if cursor.fetchone()['count'] == 0:
            default_business_types = [
                ('Buyer', 1),
                ('Agent', 2),
                ('Garment Factory', 3)
            ]
            cursor.executemany("""
            INSERT INTO customer_business_types (name, display_order) VALUES (?, ?)
            """, default_business_types)
        
        # Initialize or update countries list
        default_countries = [
            ('India', 1),
            ('China', 2),
            ('United States', 3),
            ('Indonesia', 4),
            ('Pakistan', 5),
            ('Nigeria', 6),
            ('Brazil', 7),
            ('Bangladesh', 8),
            ('Russia', 9),
            ('Ethiopia', 10),
            ('Mexico', 11),
            ('Japan', 12),
            ('Egypt', 13),
            ('Philippines', 14),
            ('DR Congo', 15),
            ('Vietnam', 16),
            ('Iran', 17),
            ('Turkey', 18),
            ('Germany', 19),
            ('Thailand', 20),
            ('Tanzania', 21),
            ('United Kingdom', 22),
            ('France', 23),
            ('South Africa', 24),
            ('Italy', 25),
            ('Kenya', 26),
            ('Myanmar', 27),
            ('Colombia', 28),
            ('South Korea', 29),
            ('Sudan', 30),
            ('Uganda', 31),
            ('Spain', 32),
            ('Algeria', 33),
            ('Iraq', 34),
            ('Argentina', 35),
            ('Afghanistan', 36),
            ('Yemen', 37),
            ('Canada', 38),
            ('Angola', 39),
            ('Ukraine', 40),
            ('Morocco', 41),
            ('Poland', 42),
            ('Uzbekistan', 43),
            ('Malaysia', 44),
            ('Mozambique', 45),
            ('Ghana', 46),
            ('Peru', 47),
            ('Saudi Arabia', 48),
            ('Madagascar', 49),
            ('Côte d\'Ivoire', 50),
            ('Cameroon', 51),
            ('Nepal', 52),
            ('Venezuela', 53),
            ('Niger', 54),
            ('Australia', 55),
            ('North Korea', 56),
            ('Syria', 57),
            ('Mali', 58),
            ('Burkina Faso', 59),
            ('Sri Lanka', 60),
            ('Malawi', 61),
            ('Zambia', 62),
            ('Chad', 63),
            ('Kazakhstan', 64),
            ('Chile', 65),
            ('Somalia', 66),
            ('Senegal', 67),
            ('Romania', 68),
            ('Guatemala', 69),
            ('Netherlands', 70),
            ('Ecuador', 71),
            ('Cambodia', 72),
            ('Zimbabwe', 73),
            ('Guinea', 74),
            ('Benin', 75),
            ('Rwanda', 76),
            ('Burundi', 77),
            ('Bolivia', 78),
            ('Tunisia', 79),
            ('South Sudan', 80),
            ('Haiti', 81),
            ('Belgium', 82),
            ('Jordan', 83),
            ('Dominican Republic', 84),
            ('United Arab Emirates', 85),
            ('Honduras', 86),
            ('Cuba', 87),
            ('Tajikistan', 88),
            ('Papua New Guinea', 89),
            ('Sweden', 90),
            ('Czech Republic (Czechia)', 91),
            ('Portugal', 92),
            ('Azerbaijan', 93),
            ('Greece', 94),
            ('Togo', 95),
            ('Hungary', 96),
            ('Israel', 97),
            ('Austria', 98),
            ('Belarus', 99),
            ('Switzerland', 100),
            ('Sierra Leone', 101),
            ('Laos', 102),
            ('Turkmenistan', 103),
            ('Libya', 104),
            ('Kyrgyzstan', 105),
            ('Paraguay', 106),
            ('Nicaragua', 107),
            ('Bulgaria', 108),
            ('Serbia', 109),
            ('Congo', 110),
            ('El Salvador', 111),
            ('Denmark', 112),
            ('Singapore', 113),
            ('Lebanon', 114),
            ('Liberia', 115),
            ('Finland', 116),
            ('Norway', 117),
            ('State of Palestine', 118),
            ('Central African Republic', 119),
            ('Oman', 120),
            ('Slovakia', 121),
            ('Mauritania', 122),
            ('Ireland', 123),
            ('New Zealand', 124),
            ('Costa Rica', 125),
            ('Kuwait', 126),
            ('Panama', 127),
            ('Croatia', 128),
            ('Georgia', 129),
            ('Eritrea', 130),
            ('Mongolia', 131),
            ('Uruguay', 132),
            ('Bosnia and Herzegovina', 133),
            ('Qatar', 134),
            ('Namibia', 135),
            ('Moldova', 136),
            ('Armenia', 137),
            ('Jamaica', 138),
            ('Lithuania', 139),
            ('Gambia', 140),
            ('Albania', 141),
            ('Gabon', 142),
            ('Botswana', 143),
            ('Lesotho', 144),
            ('Guinea-Bissau', 145),
            ('Slovenia', 146),
            ('Equatorial Guinea', 147),
            ('Latvia', 148),
            ('North Macedonia', 149),
            ('Bahrain', 150),
            ('Trinidad and Tobago', 151),
            ('Timor-Leste', 152),
            ('Cyprus', 153),
            ('Estonia', 154),
            ('Mauritius', 155),
            ('Eswatini', 156),
            ('Djibouti', 157),
            ('Fiji', 158),
            ('Comoros', 159),
            ('Solomon Islands', 160),
            ('Guyana', 161),
            ('Bhutan', 162),
            ('Luxembourg', 163),
            ('Suriname', 164),
            ('Montenegro', 165),
            ('Malta', 166),
            ('Maldives', 167),
            ('Cabo Verde', 168),
            ('Brunei', 169),
            ('Belize', 170),
            ('Bahamas', 171),
            ('Iceland', 172),
            ('Vanuatu', 173),
            ('Barbados', 174),
            ('Sao Tome & Principe', 175),
            ('Samoa', 176),
            ('Saint Lucia', 177),
            ('Kiribati', 178),
            ('Seychelles', 179),
            ('Grenada', 180),
            ('Micronesia', 181),
            ('Tonga', 182),
            ('St. Vincent & Grenadines', 183),
            ('Antigua and Barbuda', 184),
            ('Andorra', 185),
            ('Dominica', 186),
            ('Saint Kitts & Nevis', 187),
            ('Liechtenstein', 188),
            ('Monaco', 189),
            ('Marshall Islands', 190),
            ('San Marino', 191),
            ('Palau', 192),
            ('Nauru', 193),
            ('Tuvalu', 194),
            ('Holy See', 195)
        ]
        # Replace all countries with the new list
        cursor.execute("DELETE FROM countries")
        cursor.executemany("""
            INSERT INTO countries (name, display_order) VALUES (?, ?)
        """, default_countries)
        connection.commit()
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()


def insert_customer(name: str, email_suffix: str, country: str = None, website: str = None, remark: str = None, attachments: str = None, company_name: str = None, tel: str = None, source: str = None, address: str = None, business_type: str = None) -> int:
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute(
        "INSERT INTO customers (name, email_suffix, country, website, remark, attachments, company_name, tel, source, address, business_type) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (name, email_suffix, country, website, remark, attachments, company_name, tel, source, address, business_type)
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
        SELECT id, name, email_suffix, country, website, remark, attachments, company_name, tel, source, address, business_type, created_at
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

        # Handle new columns that may not exist in old databases
        country = None
        website = None
        remark = None
        attachments = None
        try:
            country = row['country'] if row['country'] else None
        except (KeyError, IndexError):
            pass
        try:
            website = row['website'] if row['website'] else None
        except (KeyError, IndexError):
            pass
        try:
            remark = row['remark'] if row['remark'] else None
        except (KeyError, IndexError):
            pass
        try:
            attachments = row['attachments'] if row['attachments'] else None
        except (KeyError, IndexError):
            pass
        try:
            company_name = row['company_name'] if row['company_name'] else None
        except (KeyError, IndexError):
            company_name = None
        try:
            tel = row['tel'] if row['tel'] else None
        except (KeyError, IndexError):
            tel = None
        try:
            source = row['source'] if row['source'] else None
        except (KeyError, IndexError):
            source = None
        try:
            address = row['address'] if row['address'] else None
        except (KeyError, IndexError):
            address = None
        try:
            business_type = row['business_type'] if row['business_type'] else None
        except (KeyError, IndexError):
            business_type = None

        customers.append({
            'id': row['id'],
            'name': row['name'],
            'email_suffix': row['email_suffix'],
            'country': country,
            'website': website,
            'source': source,
            'remark': remark,
            'attachments': attachments,
            'company_name': company_name,
            'tel': tel,
            'address': address,
            'business_type': business_type,
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
        attachments = email.get('attachments') or []
        try:
            attachments_json = json.dumps(attachments)
        except (TypeError, ValueError):
            attachments_json = '[]'
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
            attachments_json,
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
            attachments,
            fetched_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(provider, email_uid) DO UPDATE SET
            subject = excluded.subject,
            from_addr = excluded.from_addr,
            to_addr = excluded.to_addr,
            date = excluded.date,
            preview = excluded.preview,
            plain_body = excluded.plain_body,
            html_body = excluded.html_body,
            sequence = excluded.sequence,
            attachments = excluded.attachments,
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


def save_oauth_token(provider: str, creds: Credentials):
    """Save OAuth token to database"""
    connection = get_db_connection()
    cursor = connection.cursor()
    token_dict = {
        'token': creds.token,
        'refresh_token': creds.refresh_token,
        'token_uri': creds.token_uri,
        'client_id': creds.client_id,
        'client_secret': creds.client_secret,
        'scopes': json.dumps(creds.scopes) if creds.scopes else '[]'
    }
    cursor.execute("""
        INSERT INTO oauth_tokens (provider, token, refresh_token, token_uri, client_id, client_secret, scopes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(provider) DO UPDATE SET
            token = excluded.token,
            refresh_token = excluded.refresh_token,
            token_uri = excluded.token_uri,
            client_id = excluded.client_id,
            client_secret = excluded.client_secret,
            scopes = excluded.scopes,
            updated_at = datetime('now')
    """, (provider, token_dict['token'], token_dict['refresh_token'], token_dict['token_uri'],
          token_dict['client_id'], token_dict['client_secret'], token_dict['scopes']))
    connection.commit()
    cursor.close()
    connection.close()


def load_oauth_token(provider: str) -> Optional[Credentials]:
    """Load OAuth token from database"""
    connection = get_db_connection()
    cursor = connection.cursor()
    cursor.execute("SELECT * FROM oauth_tokens WHERE provider = ?", (provider,))
    row = cursor.fetchone()
    cursor.close()
    connection.close()
    
    if not row:
        return None
    
    try:
        scopes = json.loads(row['scopes']) if row['scopes'] else []
        creds = Credentials(
            token=row['token'],
            refresh_token=row['refresh_token'],
            token_uri=row['token_uri'] or 'https://oauth2.googleapis.com/token',
            client_id=row['client_id'],
            client_secret=row['client_secret'],
            scopes=scopes
        )
        
        # Refresh token if expired
        if creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                save_oauth_token(provider, creds)
            except Exception as refresh_error:
                error_str = str(refresh_error)
                # If refresh token is invalid/revoked, clear the token from database
                if 'invalid_grant' in error_str.lower() or 'token has been expired' in error_str.lower() or 'revoked' in error_str.lower():
                    print(f"OAuth token refresh failed (token expired/revoked): {refresh_error}")
                    # Clear the invalid token from database
                    connection = get_db_connection()
                    cursor = connection.cursor()
                    cursor.execute("DELETE FROM oauth_tokens WHERE provider = ?", (provider,))
                    connection.commit()
                    cursor.close()
                    connection.close()
                raise refresh_error
        
        return creds
    except Exception as e:
        print(f"Error loading OAuth token: {e}")
        return None


def fetch_gmail_api(limit=50, days_back=1):
    """Fetch emails from Gmail using Gmail API"""
    emails = []
    today = datetime.now().date()
    lookback_days = max(0, days_back)
    allowed_dates = {today - timedelta(days=offset) for offset in range(lookback_days + 1)}
    
    creds = load_oauth_token('gmail')
    if not creds:
        return {'error': 'Gmail OAuth not configured. Please authenticate first.', 'needs_auth': True}
    
    try:
        service = build('gmail', 'v1', credentials=creds)
        
        # Calculate date range for query
        oldest_date = min(allowed_dates)
        query = f'after:{oldest_date.strftime("%Y/%m/%d")}'
        
        # List messages
        results = service.users().messages().list(
            userId='me',
            q=query,
            maxResults=min(limit, 500)
        ).execute()
        
        messages = results.get('messages', [])
        
        for msg_item in messages[:limit]:
            try:
                msg = service.users().messages().get(
                    userId='me',
                    id=msg_item['id'],
                    format='full'
                ).execute()
                
                payload = msg['payload']
                headers = payload.get('headers', [])
                
                # Extract headers
                subject = next((h['value'] for h in headers if h['name'] == 'Subject'), 'No Subject')
                from_addr = next((h['value'] for h in headers if h['name'] == 'From'), '')
                to_addr = next((h['value'] for h in headers if h['name'] == 'To'), '')
                date_str = next((h['value'] for h in headers if h['name'] == 'Date'), '')
                
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
                
                if not date_obj or date_obj.date() not in allowed_dates:
                    continue
                
                # Extract body and attachments
                body_plain = ''
                body_html = ''
                attachments = []
                
                def extract_parts(parts):
                    nonlocal body_plain, body_html, attachments
                    for part in parts:
                        mime_type = part.get('mimeType', '')
                        filename = part.get('filename', '')
                        body_data = part.get('body', {})
                        attachment_id = body_data.get('attachmentId')
                        
                        # Check if it's an attachment
                        if attachment_id:
                            try:
                                att_data = service.users().messages().attachments().get(
                                    userId='me',
                                    messageId=msg_item['id'],
                                    id=attachment_id
                                ).execute()
                                
                                # Gmail API returns base64url encoded data
                                file_data = base64.urlsafe_b64decode(att_data['data'])
                                
                                attachments.append({
                                    'filename': filename or f'attachment_{len(attachments) + 1}',
                                    'content_type': mime_type or 'application/octet-stream',
                                    'size': len(file_data),
                                    'data': base64.b64encode(file_data).decode('ascii')
                                })
                            except Exception as att_exc:
                                print(f"Error fetching attachment: {att_exc}")
                                continue
                        elif filename and mime_type not in ['text/plain', 'text/html']:
                            # Inline attachment without attachmentId - try to get from body data
                            data = body_data.get('data', '')
                            if data:
                                try:
                                    file_data = base64.urlsafe_b64decode(data)
                                    attachments.append({
                                        'filename': filename,
                                        'content_type': mime_type or 'application/octet-stream',
                                        'size': len(file_data),
                                        'data': base64.b64encode(file_data).decode('ascii')
                                    })
                                except Exception:
                                    pass
                        else:
                            # Extract body text
                            data = body_data.get('data', '')
                            if data:
                                try:
                                    decoded = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                                    if mime_type == 'text/plain' and not body_plain:
                                        body_plain = decoded
                                    elif mime_type == 'text/html' and not body_html:
                                        body_html = decoded
                                except Exception:
                                    pass
                        
                        # Recursively process nested parts
                        if 'parts' in part:
                            extract_parts(part['parts'])
                
                # Process payload
                if 'parts' in payload:
                    extract_parts(payload['parts'])
                else:
                    # Single part message
                    mime_type = payload.get('mimeType', '')
                    body_data = payload.get('body', {})
                    data = body_data.get('data', '')
                    if data:
                        try:
                            decoded = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
                            if mime_type == 'text/html':
                                body_html = decoded
                            else:
                                body_plain = decoded
                        except Exception:
                            pass
                
                preview_source = (body_plain or strip_html_tags(body_html) or '').strip()
                preview = preview_source[:500] + '...' if len(preview_source) > 500 else preview_source
                sequence_code = build_sequence_code(from_addr, date_obj)
                
                emails.append({
                    'id': msg_item['id'],
                    'subject': subject,
                    'from': from_addr,
                    'to': to_addr,
                    'date': date_formatted,
                    'preview': preview,
                    'plain_body': body_plain,
                    'html_body': body_html,
                    'sequence': sequence_code,
                    'attachments': attachments
                })
            except Exception as e:
                print(f"Error processing email {msg_item.get('id', 'unknown')}: {str(e)}")
                continue
        
        return {'emails': emails, 'count': len(emails)}
    except HttpError as e:
        error_str = str(e)
        # Check if it's an authentication error
        if 'invalid_grant' in error_str.lower() or 'token has been expired' in error_str.lower() or 'revoked' in error_str.lower():
            # Clear the invalid token
            try:
                connection = get_db_connection()
                cursor = connection.cursor()
                cursor.execute("DELETE FROM oauth_tokens WHERE provider = ?", ('gmail',))
                connection.commit()
                cursor.close()
                connection.close()
            except Exception:
                pass
            return {'error': 'Gmail OAuth token has expired or been revoked. Please re-authenticate.', 'needs_auth': True}
        return {'error': f'Gmail API error: {error_str}'}
    except Exception as e:
        error_str = str(e)
        # Check if it's an authentication error
        if 'invalid_grant' in error_str.lower() or 'token has been expired' in error_str.lower() or 'revoked' in error_str.lower():
            # Clear the invalid token
            try:
                connection = get_db_connection()
                cursor = connection.cursor()
                cursor.execute("DELETE FROM oauth_tokens WHERE provider = ?", ('gmail',))
                connection.commit()
                cursor.close()
                connection.close()
            except Exception:
                pass
            return {'error': 'Gmail OAuth token has expired or been revoked. Please re-authenticate.', 'needs_auth': True}
        return {'error': f'Error fetching Gmail: {error_str}'}


def build_sequence_code(from_address: str, email_date: Optional[datetime] = None) -> str:
    """Construct a sequence code yyyymmdd_hhmmss_<two letters>_<domain>."""
    email_datetime = email_date or datetime.now()
    sequence_date = email_datetime.strftime('%Y%m%d_%H%M%S')
    
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


def fetch_emails(imap_server, port, username, password, use_ssl=True, use_tls=False, limit=50, days_back=0):
    """Fetch emails from IMAP server"""
    emails = []
    today = datetime.now().date()
    lookback_days = max(0, days_back)
    allowed_dates = {today - timedelta(days=offset) for offset in range(lookback_days + 1)}

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
        
        # Search for recent emails - fetch all emails in date range (with and without attachments)
        oldest_date = min(allowed_dates)
        since_clause = f'(SINCE { _imap_date(oldest_date) })'
        status, messages = mail.search(None, since_clause)
        email_ids = []
        if status == 'OK' and messages and len(messages) > 0:
            email_ids = [seq_id for seq_id in messages[0].split() if seq_id and seq_id.strip()]
        
        # Get the most recent emails (limit)
        if email_ids:
            email_ids = sorted(email_ids, key=lambda eid: int(eid))
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

                if not date_obj or date_obj.date() not in allowed_dates:
                    continue

                # Get email body (plain and html) and attachments
                body_plain = ''
                body_html = ''
                attachments = []

                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition") or "").lower()
                        filename = part.get_filename()

                        # More thorough attachment detection
                        is_attachment = (
                            'attachment' in content_disposition or
                            bool(filename) or
                            (content_type not in ['text/plain', 'text/html', 'multipart/alternative', 'multipart/related', 'multipart/mixed'] and 
                             'inline' not in content_disposition)
                        )

                        # Only extract body from non-attachment text parts
                        if not is_attachment:
                            if content_type == "text/plain" and not body_plain:
                                try:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        body_plain = payload.decode('utf-8', errors='ignore')
                                except Exception:
                                    pass
                            elif content_type == "text/html" and not body_html:
                                try:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        body_html = payload.decode('utf-8', errors='ignore')
                                except Exception:
                                    pass

                        # Extract attachments
                        if is_attachment:
                            try:
                                payload = part.get_payload(decode=True)
                                if payload is None:
                                    payload = b''
                                elif not isinstance(payload, bytes):
                                    payload = str(payload).encode('utf-8')
                                
                                decoded_filename = decode_mime_words(filename) if filename else f'attachment_{len(attachments) + 1}'
                                
                                # Skip empty attachments
                                if len(payload) > 0:
                                    attachments.append({
                                        'filename': decoded_filename,
                                        'content_type': content_type or 'application/octet-stream',
                                        'size': len(payload),
                                        'data': base64.b64encode(payload).decode('ascii')
                                    })
                            except Exception as att_exc:
                                print(f"Error extracting attachment: {att_exc}")
                                continue
                else:
                    content_type = msg.get_content_type()
                    filename = msg.get_filename()
                    is_attachment = bool(filename)
                    try:
                        payload = msg.get_payload(decode=True)
                        decoded = payload.decode('utf-8', errors='ignore') if isinstance(payload, bytes) else str(payload)
                    except Exception:
                        payload = b''
                        decoded = str(msg.get_payload())

                    if is_attachment:
                        attachments.append({
                            'filename': decode_mime_words(filename) if filename else 'attachment',
                            'content_type': content_type,
                            'size': len(payload or b''),
                            'data': base64.b64encode(payload or b'').decode('ascii')
                        })
                    elif content_type == "text/html":
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
                    'sequence': sequence_code,
                    'attachments': attachments
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


@app.route('/api/gmail-auth', methods=['GET'])
def gmail_auth():
    """Start Gmail OAuth flow"""
    data = request.args
    client_id = (data.get('client_id') or GMAIL_OAUTH_CONFIG.get('client_id') or '').strip()
    client_secret = (data.get('client_secret') or GMAIL_OAUTH_CONFIG.get('client_secret') or '').strip()
    redirect_uri = (data.get('redirect_uri') or GMAIL_OAUTH_CONFIG.get('redirect_uri') or 'http://localhost:5000/oauth2callback').strip()
    
    if not client_id or not client_secret:
        return jsonify({'error': 'Gmail OAuth client_id and client_secret are required. Please enter them in the Settings page.'}), 400
    
    # Validate Client ID format
    if not client_id.endswith('.apps.googleusercontent.com'):
        return jsonify({'error': f'Invalid Client ID format. Should end with .apps.googleusercontent.com. Got: {client_id[:50]}...'}), 400
    
    # Validate Client Secret format
    if not client_secret.startswith('GOCSPX-'):
        return jsonify({'error': f'Invalid Client Secret format. Should start with GOCSPX-. Please check your Google Cloud Console.'}), 400
    
    try:
        flow = Flow.from_client_config(
            {
                "web": {
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": [redirect_uri]
                }
            },
            scopes=GMAIL_OAUTH_CONFIG['scopes'],
            redirect_uri=redirect_uri
        )
        
        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true',
            prompt='consent'
        )
        
        # Store state in both session and database for reliability
        session['oauth_state'] = state
        session['oauth_client_id'] = client_id
        session['oauth_client_secret'] = client_secret
        session['oauth_redirect_uri'] = redirect_uri
        
        # Also store in database as backup (in case session is lost)
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO oauth_states (state, client_id, client_secret, redirect_uri, created_at)
                VALUES (?, ?, ?, ?, datetime('now'))
            """, (state, client_id, client_secret, redirect_uri))
            connection.commit()
            cursor.close()
            connection.close()
        except Exception as e:
            print(f"Warning: Could not store OAuth state in database: {e}")
        
        return jsonify({'auth_url': authorization_url})
    except Exception as e:
        error_msg = str(e)
        # Provide more helpful error messages
        if 'invalid_client' in error_msg.lower():
            return jsonify({
                'error': f'Invalid Client ID or Client Secret. Please verify:\n'
                        f'1. Client ID: {client_id[:30]}...\n'
                        f'2. Client Secret: {client_secret[:10]}...\n'
                        f'3. Make sure they match the credentials from Google Cloud Console\n'
                        f'4. Ensure the OAuth client type is "Web application" (not Desktop)'
            }), 400
        return jsonify({'error': f'OAuth setup error: {error_msg}'}), 500


@app.route('/oauth2callback')
def oauth2callback():
    """OAuth 2.0 callback handler"""
    code = request.args.get('code')
    state = request.args.get('state')
    
    if not code:
        return '<html><body><h1>Authentication Failed</h1><p>No authorization code received.</p><script>setTimeout(() => window.close(), 3000);</script></body></html>', 400
    
    if not state:
        return '<html><body><h1>Authentication Failed</h1><p>No state parameter received.</p><script>setTimeout(() => window.close(), 3000);</script></body></html>', 400
    
    # Try to get credentials from session first
    client_id = session.get('oauth_client_id')
    client_secret = session.get('oauth_client_secret')
    redirect_uri = session.get('oauth_redirect_uri')
    session_state = session.get('oauth_state')
    
    # If session state doesn't match, try to get from database
    if state != session_state:
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("""
                SELECT client_id, client_secret, redirect_uri 
                FROM oauth_states 
                WHERE state = ? AND datetime(created_at) > datetime('now', '-10 minutes')
            """, (state,))
            row = cursor.fetchone()
            if row:
                client_id = row['client_id']
                client_secret = row['client_secret']
                redirect_uri = row['redirect_uri']
                # Delete used state
                cursor.execute("DELETE FROM oauth_states WHERE state = ?", (state,))
                connection.commit()
            cursor.close()
            connection.close()
        except Exception as e:
            print(f"Error checking OAuth state in database: {e}")
    
    # Final validation
    if state != session_state and not client_id:
        return '<html><body><h1>Authentication Failed</h1><p>Invalid state parameter. Please try authenticating again.</p><script>setTimeout(() => window.close(), 3000);</script></body></html>', 400
    
    if not client_id or not client_secret:
        return '<html><body><h1>Authentication Failed</h1><p>OAuth credentials not found. Please try authenticating again.</p><script>setTimeout(() => window.close(), 3000);</script></body></html>', 400
    
    if not redirect_uri:
        redirect_uri = GMAIL_OAUTH_CONFIG.get('redirect_uri', 'http://localhost:5000/oauth2callback')
    
    try:
        
        flow = Flow.from_client_config(
            {
                "web": {
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": [redirect_uri]
                }
            },
            scopes=GMAIL_OAUTH_CONFIG['scopes'],
            redirect_uri=redirect_uri,
            state=state
        )
        
        flow.fetch_token(code=code)
        creds = flow.credentials
        
        # Save credentials
        save_oauth_token('gmail', creds)
        
        return '<html><body><h1>Authentication Successful!</h1><p>You can close this window and return to the app.</p><script>setTimeout(() => window.close(), 2000);</script></body></html>'
    except Exception as e:
        error_msg = str(e)
        error_html = '<html><body style="font-family: Arial, sans-serif; padding: 20px;"><h1 style="color: #dc3545;">Authentication Error</h1>'
        
        if 'invalid_client' in error_msg.lower():
            error_html += f'''
            <div style="background-color: #f8d7da; border: 1px solid #dc3545; padding: 15px; border-radius: 4px; margin: 15px 0;">
                <h2 style="color: #721c24; margin-top: 0;">Invalid Client ID or Client Secret</h2>
                <p style="color: #721c24;"><strong>This error means your Client ID and Client Secret don't match.</strong></p>
                <ol style="color: #721c24;">
                    <li>Go to <a href="https://console.cloud.google.com/apis/credentials" target="_blank">Google Cloud Console Credentials</a></li>
                    <li>Find your OAuth client: <code style="background: #fff; padding: 2px 5px;">{client_id[:50]}...</code></li>
                    <li>Click on it to view details</li>
                    <li>Copy the <strong>Client ID</strong> and <strong>Client Secret</strong> again</li>
                    <li>Make sure you're copying from a <strong>"Web application"</strong> type client (not Desktop)</li>
                    <li>Paste them in the Settings page and try again</li>
                </ol>
                <p style="color: #721c24; margin-bottom: 0;"><strong>Note:</strong> Client Secret starts with "GOCSPX-" and you can only see it once when you create the client.</p>
            </div>
            '''
        else:
            error_html += f'<p style="color: #721c24;">{error_msg}</p>'
        
        error_html += '<p><button onclick="window.close()">Close Window</button></p>'
        error_html += '<script>setTimeout(() => window.close(), 10000);</script></body></html>'
        return error_html, 500


@app.route('/api/gmail-status', methods=['GET'])
def gmail_status():
    """Check Gmail OAuth authentication status"""
    creds = load_oauth_token('gmail')
    if creds:
        try:
            # Try to verify token is valid
            service = build('gmail', 'v1', credentials=creds)
            profile = service.users().getProfile(userId='me').execute()
            return jsonify({
                'authenticated': True,
                'email': profile.get('emailAddress', ''),
                'messages_total': profile.get('messagesTotal', 0)
            })
        except Exception as e:
            return jsonify({
                'authenticated': False,
                'error': f'Token invalid: {str(e)}',
                'needs_auth': True
            })
    return jsonify({
        'authenticated': False,
        'needs_auth': True,
        'message': 'Gmail OAuth not configured. Please set up OAuth credentials in Settings.'
    })


@app.route('/api/fetch-gmail', methods=['POST'])
def fetch_gmail():
    """Fetch emails from Gmail using Gmail API"""
    data = request.json or {}
    limit = data.get('limit', 50)
    days_back = data.get('days_back', 1)
    try:
        limit = max(1, int(limit))
    except (TypeError, ValueError):
        limit = 50
    try:
        days_back = max(0, int(days_back))
    except (TypeError, ValueError):
        days_back = 1
    
    # Check if OAuth is configured
    creds = load_oauth_token('gmail')
    if not creds:
        return jsonify({
            'error': 'Gmail OAuth not authenticated. Please:\n1. Go to Settings\n2. Enter OAuth Client ID and Client Secret\n3. Click "Authenticate Gmail"\n4. Complete the OAuth flow',
            'needs_auth': True,
            'setup_url': 'https://console.cloud.google.com/'
        }), 401
    
    # Fetch emails for requested time window
    result = fetch_gmail_api(limit=limit, days_back=days_back)
    
    if 'error' in result:
        # If authentication is needed, return 401 status
        if result.get('needs_auth'):
            return jsonify(result), 401
        return jsonify(result), 500
    
    # Save emails to SQL database
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
        limit,
        days_back=1
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
        limit,
        days_back=1
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
    # Support both email_address (new) and email_suffix (old) for backward compatibility
    email_address = (data.get('email_address') or '').strip()
    suffix = (data.get('email_suffix') or '').strip()
    country = (data.get('country') or '').strip() or None
    website = (data.get('website') or '').strip() or None
    source = (data.get('source') or '').strip() or None
    address = (data.get('address') or '').strip() or None
    business_type = (data.get('business_type') or '').strip() or None
    remark = (data.get('remark') or '').strip() or None
    attachments = data.get('attachments') or None
    company_name = (data.get('company_name') or '').strip() or None
    tel = (data.get('tel') or '').strip() or None

    if not name:
        return jsonify({'error': 'Customer name is required'}), 400
    
    # If email_address is provided, validate and save full email address
    if email_address:
        # Validate email format
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        if not re.match(email_pattern, email_address):
            return jsonify({'error': 'Invalid email address format'}), 400
        # Save full email address (prefix + suffix)
        full_email = email_address
    elif suffix:
        # Backward compatibility: handle old email_suffix format
        if '@' in suffix:
            # If it already contains @, treat as full email
            full_email = suffix
        else:
            # Old format: just domain, convert to @domain format for backward compatibility
            if not re.match(r'^[a-zA-Z0-9._-]+\.[a-zA-Z]{2,}$', suffix):
                return jsonify({'error': 'Invalid email suffix format'}), 400
            full_email = '@' + suffix
    else:
        return jsonify({'error': 'Email address is required'}), 400

    try:
        customer_id = insert_customer(name, full_email, country, website, remark, attachments, company_name, tel, source, address, business_type)
        return jsonify({
            'id': customer_id,
            'name': name,
            'email_suffix': full_email,
            'country': country,
            'tel': tel,
            'website': website,
            'source': source,
            'remark': remark,
            'attachments': attachments,
            'company_name': company_name,
            'address': address,
            'business_type': business_type
        }), 201
    except Exception as exc:
        return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/customers/<int:customer_id>', methods=['PUT', 'DELETE'])
def update_or_delete_customer(customer_id):
    """Update or delete a customer by ID"""
    if request.method == 'PUT':
        try:
            data = request.json or {}
            name = (data.get('name') or '').strip()
            email_address = (data.get('email_address') or '').strip()
            country = (data.get('country') or '').strip() or None
            website = (data.get('website') or '').strip() or None
            source = (data.get('source') or '').strip() or None
            remark = (data.get('remark') or '').strip() or None
            attachments = data.get('attachments') or None
            company_name = (data.get('company_name') or '').strip() or None
            tel = (data.get('tel') or '').strip() or None
            address = (data.get('address') or '').strip() or None
            business_type = (data.get('business_type') or '').strip() or None
            
            if not name:
                return jsonify({'error': 'Customer name is required'}), 400
            
            if email_address:
                email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                if not re.match(email_pattern, email_address):
                    return jsonify({'error': 'Invalid email address format'}), 400
                full_email = email_address
            else:
                return jsonify({'error': 'Email address is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("""
                UPDATE customers 
                SET name = ?, email_suffix = ?, country = ?, website = ?, source = ?, remark = ?, attachments = ?, company_name = ?, tel = ?, address = ?, business_type = ?
                WHERE id = ?
            """, (name, full_email, country, website, source, remark, attachments, company_name, tel, address, business_type, customer_id))
            connection.commit()
            updated = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if updated:
                return jsonify({
                    'id': customer_id,
                    'name': name,
                    'email_suffix': full_email,
                    'country': country,
                    'tel': tel,
                    'website': website,
                    'source': source,
                    'remark': remark,
                    'attachments': attachments,
                    'company_name': company_name,
                    'address': address,
                    'business_type': business_type,
                    'status': 'updated'
                })
            else:
                return jsonify({'error': 'Customer not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM customers WHERE id = ?", (customer_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': customer_id})
            else:
                return jsonify({'error': 'Customer not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/emails', methods=['GET', 'POST'])
def handle_emails():
    if request.method == 'GET':
        provider = (request.args.get('provider') or '').strip().lower()
        if not provider:
            return jsonify({'error': 'Provider query parameter is required'}), 400
        
        # Optional days parameter (number of days back to include)
        days_param = request.args.get('days')
        default_days = 1
        if days_param is None or days_param == '':
            days = default_days
        else:
            try:
                days = max(0, int(days_param))
            except (TypeError, ValueError):
                return jsonify({'error': 'Invalid days parameter'}), 400

        # Build SQL query with optional date filter
        connection = get_db_connection()
        cursor = connection.cursor()
        query = """
            SELECT provider, email_uid, subject, from_addr, to_addr, date, preview, plain_body, html_body, sequence, attachments, fetched_at
            FROM emails
            WHERE provider = ?
        """
        params = [provider]
        if days > 0:
            query += " AND datetime(fetched_at) >= datetime('now', ?)"
            params.append(f'-{days} days')
        query += " ORDER BY datetime(date) DESC, datetime(fetched_at) DESC"
        cursor.execute(query, tuple(params))
        rows = cursor.fetchall()
        cursor.close()
        connection.close()

        emails = []
        for row in rows:
            attachments = []
            try:
                attachments = json.loads(row[10] or '[]')
            except (TypeError, ValueError):
                attachments = []
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
                'attachments': attachments,
                'fetched_at': row[11]
            })

        return jsonify({'provider': provider, 'emails': emails})
    
    elif request.method == 'POST':
        data = request.json or {}
        provider = (data.get('provider') or '').strip().lower()
        emails = data.get('emails') or []

        if not provider:
            return jsonify({'error': 'Provider is required'}), 400

        if not isinstance(emails, list):
            return jsonify({'error': 'Emails must be a list'}), 400

        save_emails(provider, emails)
        return jsonify({'status': 'saved', 'count': len(emails)})


@app.route('/api/version', methods=['GET'])
def get_version():
    """Get application version"""
    return jsonify({'version': app.config['VERSION']})


@app.route('/api/tasks', methods=['GET', 'POST'])
def handle_tasks():
    """Handle task creation and retrieval"""
    if request.method == 'GET':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Always try to add email column if it doesn't exist (idempotent)
            try:
                cursor.execute("ALTER TABLE tasks ADD COLUMN email TEXT")
                connection.commit()
                # Force schema refresh by querying table info
                cursor.execute("PRAGMA table_info(tasks)")
            except sqlite3.OperationalError:
                connection.rollback()  # Column already exists, rollback any partial changes
                pass
            
            # Try the query with email column first
            try:
                cursor.execute("""
                    SELECT 
                        t.id,
                        t.sequence,
                        t.customer,
                        t.email,
                        t.catalogue,
                        t.template,
                        t.attachments,
                        t.deadline,
                        t.created_at,
                        t.updated_at,
                        (SELECT company_name FROM customers c 
                         WHERE (t.email IS NOT NULL AND c.email_suffix = t.email) 
                            OR (t.customer IS NOT NULL AND c.name = t.customer)
                         ORDER BY CASE WHEN t.email IS NOT NULL AND c.email_suffix = t.email THEN 1 ELSE 2 END, c.id
                         LIMIT 1) AS company_name,
                        (SELECT source FROM customers c 
                         WHERE (t.email IS NOT NULL AND c.email_suffix = t.email) 
                            OR (t.customer IS NOT NULL AND c.name = t.customer)
                         ORDER BY CASE WHEN t.email IS NOT NULL AND c.email_suffix = t.email THEN 1 ELSE 2 END, c.id
                         LIMIT 1) AS customer_source,
                        (SELECT business_type FROM customers c 
                         WHERE (t.email IS NOT NULL AND c.email_suffix = t.email) 
                            OR (t.customer IS NOT NULL AND c.name = t.customer)
                         ORDER BY CASE WHEN t.email IS NOT NULL AND c.email_suffix = t.email THEN 1 ELSE 2 END, c.id
                         LIMIT 1) AS customer_business_type
                    FROM tasks t
                    ORDER BY datetime(t.created_at) DESC
                """)
            except sqlite3.OperationalError as e:
                # If email column still doesn't exist, use fallback query
                if 'no such column' in str(e).lower() and 'email' in str(e).lower():
                    cursor.execute("""
                        SELECT 
                            t.id,
                            t.sequence,
                            t.customer,
                            NULL as email,
                            t.catalogue,
                            t.template,
                            t.attachments,
                            t.deadline,
                            t.created_at,
                            t.updated_at,
                            (SELECT company_name FROM customers c 
                             WHERE c.name = t.customer
                             ORDER BY c.id
                             LIMIT 1) AS company_name,
                            (SELECT source FROM customers c 
                             WHERE c.name = t.customer
                             ORDER BY c.id
                             LIMIT 1) AS customer_source,
                            (SELECT business_type FROM customers c 
                             WHERE c.name = t.customer
                             ORDER BY c.id
                             LIMIT 1) AS customer_business_type
                        FROM tasks t
                        ORDER BY datetime(t.created_at) DESC
                    """)
                else:
                    raise  # Re-raise if it's a different error
            rows = cursor.fetchall()
            cursor.close()
            connection.close()
            
            tasks = []
            for row in rows:
                attachments = []
                try:
                    attachments = json.loads(row[6] or '[]')
                except (TypeError, ValueError):
                    attachments = []
                
                tasks.append({
                    'id': row['id'],
                    'sequence': row['sequence'],
                    'customer': row['customer'],
                    'email': row['email'],
                    'catalogue': row['catalogue'],
                    'template': row['template'],
                    'attachments': attachments,
                    'deadline': row['deadline'],
                    'created_at': row['created_at'],
                    'updated_at': row['updated_at'],
                    'company_name': row['company_name'],
                    'source': row['customer_source'],
                    'business_type': row['customer_business_type']
                })
            
            return jsonify({'tasks': tasks})
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'POST':
        data = request.json or {}
        sequence = data.get('sequence', '')
        customer = data.get('customer', '')
        email = data.get('email', '')
        catalogue = data.get('catalogue', '').strip()
        template = data.get('template', '').strip()
        attachments = data.get('attachments', [])
        deadline = data.get('deadline')
        if isinstance(deadline, str):
            deadline = deadline.strip() or None
        
        if not catalogue:
            return jsonify({'error': 'Catalogue is required'}), 400
        
        if not template:
            return jsonify({'error': 'Template is required'}), 400
        
        try:
            attachments_json = json.dumps(attachments) if attachments else '[]'
            
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("""
                INSERT INTO tasks (sequence, customer, email, catalogue, template, attachments, deadline, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
            """, (sequence, customer, email, catalogue, template, attachments_json, deadline))
            connection.commit()
            task_id = cursor.lastrowid
            cursor.execute("SELECT created_at, updated_at FROM tasks WHERE id = ?", (task_id,))
            timestamps_row = cursor.fetchone()
            created_at_value = timestamps_row['created_at'] if timestamps_row else None
            updated_at_value = timestamps_row['updated_at'] if timestamps_row else None
            cursor.close()
            connection.close()
            
            return jsonify({
                'id': task_id,
                'sequence': sequence,
                'customer': customer,
                'email': email,
                'catalogue': catalogue,
                'template': template,
                'attachments': attachments,
                'deadline': deadline,
                'created_at': created_at_value,
                'updated_at': updated_at_value
            }), 201
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/tasks/<int:task_id>', methods=['PUT', 'DELETE'])
def handle_single_task(task_id):
    """Update or delete a task by ID"""
    if request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': task_id})
            else:
                return jsonify({'error': 'Task not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    # PUT - update task
    try:
        data = request.get_json() or {}
        catalogue = (data.get('catalogue') or '').strip()
        template = (data.get('template') or '').strip()
        email = (data.get('email') or '').strip()
        customer = (data.get('customer') or '').strip()
        attachments = data.get('attachments')
        deadline = (data.get('deadline') or '').strip() or None
        
        if not catalogue:
            return jsonify({'error': 'Catalogue is required'}), 400
        if not template:
            return jsonify({'error': 'Template is required'}), 400
        
        attachments_json = None
        if attachments is not None:
            try:
                attachments_json = json.dumps(attachments)
            except (TypeError, ValueError):
                return jsonify({'error': 'Invalid attachments format'}), 400
        
        connection = get_db_connection()
        cursor = connection.cursor()
        
        update_fields = [
            ('catalogue', catalogue),
            ('template', template),
            ('deadline', deadline)
        ]
        
        if email:
            update_fields.append(('email', email))
        if customer:
            update_fields.append(('customer', customer))
        if attachments_json is not None:
            update_fields.append(('attachments', attachments_json))
        
        set_clause = ', '.join([f"{field} = ?" for field, _ in update_fields])
        if set_clause:
            set_clause = f"{set_clause}, updated_at = datetime('now')"
        values = [value for _, value in update_fields]
        values.append(task_id)
        
        cursor.execute(f"UPDATE tasks SET {set_clause} WHERE id = ?", values)
        connection.commit()
        updated = cursor.rowcount > 0
        cursor.close()
        connection.close()
        
        if not updated:
            return jsonify({'error': 'Task not found'}), 404
        
        return jsonify({'status': 'updated', 'id': task_id})
    except Exception as exc:
        return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/task-types', methods=['GET', 'POST'])
def handle_task_types():
    """Handle task type retrieval and creation"""
    if request.method == 'GET':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("SELECT id, name, display_order, created_at FROM task_types ORDER BY display_order, name")
            types = [dict(row) for row in cursor.fetchall()]
            cursor.close()
            connection.close()
            return jsonify(types)
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Task type name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists
            cursor.execute("SELECT id FROM task_types WHERE name = ?", (name,))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Task type with this name already exists'}), 400
            
            # Get max display_order
            cursor.execute("SELECT MAX(display_order) as max_order FROM task_types")
            max_order = cursor.fetchone()['max_order'] or 0
            display_order = data.get('display_order', max_order + 1)
            
            cursor.execute("""
                INSERT INTO task_types (name, display_order) VALUES (?, ?)
            """, (name, display_order))
            connection.commit()
            task_type_id = cursor.lastrowid
            cursor.close()
            connection.close()
            
            return jsonify({
                'id': task_type_id,
                'name': name,
                'display_order': display_order,
                'status': 'created'
            }), 201
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/task-types/<int:type_id>', methods=['PUT', 'DELETE'])
def handle_task_type(type_id):
    """Handle task type update and deletion"""
    if request.method == 'PUT':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Task type name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists for another type
            cursor.execute("SELECT id FROM task_types WHERE name = ? AND id != ?", (name, type_id))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Task type with this name already exists'}), 400
            
            display_order = data.get('display_order', 0)
            cursor.execute("""
                UPDATE task_types SET name = ?, display_order = ? WHERE id = ?
            """, (name, display_order, type_id))
            connection.commit()
            updated = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if updated:
                return jsonify({
                    'id': type_id,
                    'name': name,
                    'display_order': display_order,
                    'status': 'updated'
                })
            else:
                return jsonify({'error': 'Task type not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM task_types WHERE id = ?", (type_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': type_id})
            else:
                return jsonify({'error': 'Task type not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/countries', methods=['GET', 'POST'])
def handle_countries():
    """Handle country retrieval and creation"""
    if request.method == 'GET':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("SELECT id, name, display_order, created_at FROM countries ORDER BY display_order, name")
            countries = [dict(row) for row in cursor.fetchall()]
            cursor.close()
            connection.close()
            return jsonify(countries)
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Country name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists
            cursor.execute("SELECT id FROM countries WHERE name = ?", (name,))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Country with this name already exists'}), 400
            
            # Get max display_order
            cursor.execute("SELECT MAX(display_order) as max_order FROM countries")
            max_order = cursor.fetchone()['max_order'] or 0
            display_order = data.get('display_order', max_order + 1)
            
            cursor.execute("""
                INSERT INTO countries (name, display_order) VALUES (?, ?)
            """, (name, display_order))
            connection.commit()
            country_id = cursor.lastrowid
            cursor.close()
            connection.close()
            
            return jsonify({
                'id': country_id,
                'name': name,
                'display_order': display_order,
                'status': 'created'
            }), 201
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/countries/<int:country_id>', methods=['PUT', 'DELETE'])
def handle_country(country_id):
    """Handle country update and deletion"""
    if request.method == 'PUT':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Country name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists for another country
            cursor.execute("SELECT id FROM countries WHERE name = ? AND id != ?", (name, country_id))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Country with this name already exists'}), 400
            
            display_order = data.get('display_order', 0)
            cursor.execute("""
                UPDATE countries SET name = ?, display_order = ? WHERE id = ?
            """, (name, display_order, country_id))
            connection.commit()
            updated = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if updated:
                return jsonify({
                    'id': country_id,
                    'name': name,
                    'display_order': display_order,
                    'status': 'updated'
                })
            else:
                return jsonify({'error': 'Country not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM countries WHERE id = ?", (country_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': country_id})
            else:
                return jsonify({'error': 'Country not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/customer-sources', methods=['GET', 'POST'])
def handle_customer_sources():
    """Handle customer source retrieval and creation"""
    if request.method == 'GET':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("SELECT id, name, display_order, created_at FROM customer_sources ORDER BY display_order, name")
            sources = [dict(row) for row in cursor.fetchall()]
            cursor.close()
            connection.close()
            return jsonify(sources)
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Customer source name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists
            cursor.execute("SELECT id FROM customer_sources WHERE name = ?", (name,))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Customer source with this name already exists'}), 400
            
            # Get max display_order
            cursor.execute("SELECT MAX(display_order) as max_order FROM customer_sources")
            max_order = cursor.fetchone()['max_order'] or 0
            display_order = data.get('display_order', max_order + 1)
            
            cursor.execute("""
                INSERT INTO customer_sources (name, display_order) VALUES (?, ?)
            """, (name, display_order))
            connection.commit()
            source_id = cursor.lastrowid
            cursor.close()
            connection.close()
            
            return jsonify({
                'id': source_id,
                'name': name,
                'display_order': display_order,
                'status': 'created'
            }), 201
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/customer-sources/<int:source_id>', methods=['PUT', 'DELETE'])
def handle_customer_source(source_id):
    """Handle customer source update and deletion"""
    if request.method == 'PUT':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Customer source name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists for another source
            cursor.execute("SELECT id FROM customer_sources WHERE name = ? AND id != ?", (name, source_id))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Customer source with this name already exists'}), 400
            
            display_order = data.get('display_order', 0)
            cursor.execute("""
                UPDATE customer_sources SET name = ?, display_order = ? WHERE id = ?
            """, (name, display_order, source_id))
            connection.commit()
            updated = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if updated:
                return jsonify({
                    'id': source_id,
                    'name': name,
                    'display_order': display_order,
                    'status': 'updated'
                })
            else:
                return jsonify({'error': 'Customer source not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM customer_sources WHERE id = ?", (source_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': source_id})
            else:
                return jsonify({'error': 'Customer source not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/customer-business-types', methods=['GET', 'POST'])
def handle_customer_business_types():
    """Handle customer business type retrieval and creation"""
    if request.method == 'GET':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("SELECT id, name, display_order, created_at FROM customer_business_types ORDER BY display_order, name")
            types = [dict(row) for row in cursor.fetchall()]
            cursor.close()
            connection.close()
            return jsonify(types)
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Customer business type name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists
            cursor.execute("SELECT id FROM customer_business_types WHERE name = ?", (name,))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Customer business type with this name already exists'}), 400
            
            # Get max display_order
            cursor.execute("SELECT MAX(display_order) as max_order FROM customer_business_types")
            max_order = cursor.fetchone()['max_order'] or 0
            display_order = data.get('display_order', max_order + 1)
            
            cursor.execute("""
                INSERT INTO customer_business_types (name, display_order) VALUES (?, ?)
            """, (name, display_order))
            connection.commit()
            type_id = cursor.lastrowid
            cursor.close()
            connection.close()
            
            return jsonify({
                'id': type_id,
                'name': name,
                'display_order': display_order,
                'status': 'created'
            }), 201
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


@app.route('/api/customer-business-types/<int:type_id>', methods=['PUT', 'DELETE'])
def handle_customer_business_type(type_id):
    """Handle customer business type update and deletion"""
    if request.method == 'PUT':
        try:
            data = request.get_json()
            name = data.get('name', '').strip()
            if not name:
                return jsonify({'error': 'Customer business type name is required'}), 400
            
            connection = get_db_connection()
            cursor = connection.cursor()
            
            # Check if name already exists for another type
            cursor.execute("SELECT id FROM customer_business_types WHERE name = ? AND id != ?", (name, type_id))
            if cursor.fetchone():
                cursor.close()
                connection.close()
                return jsonify({'error': 'Customer business type with this name already exists'}), 400
            
            display_order = data.get('display_order', 0)
            cursor.execute("""
                UPDATE customer_business_types SET name = ?, display_order = ? WHERE id = ?
            """, (name, display_order, type_id))
            connection.commit()
            updated = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if updated:
                return jsonify({
                    'id': type_id,
                    'name': name,
                    'display_order': display_order,
                    'status': 'updated'
                })
            else:
                return jsonify({'error': 'Customer business type not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500
    
    elif request.method == 'DELETE':
        try:
            connection = get_db_connection()
            cursor = connection.cursor()
            cursor.execute("DELETE FROM customer_business_types WHERE id = ?", (type_id,))
            connection.commit()
            deleted = cursor.rowcount > 0
            cursor.close()
            connection.close()
            
            if deleted:
                return jsonify({'status': 'deleted', 'id': type_id})
            else:
                return jsonify({'error': 'Customer business type not found'}), 404
        except Exception as exc:
            return jsonify({'error': f'Database error: {str(exc)}'}), 500


initialize_database()


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

