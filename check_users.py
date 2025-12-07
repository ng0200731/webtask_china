import sqlite3
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / 'mailtask.db'

conn = sqlite3.connect(str(DB_PATH))
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

cursor.execute("SELECT email, level, status FROM users WHERE email IN ('eric.brilliant@gmail.com', 'weiwu@fuchanghk.com')")
rows = cursor.fetchall()

print("User levels:")
for row in rows:
    print(f"  {row['email']}: Level {row['level']}, Status: {row['status']}")

conn.close()


