import sqlite3

db_path = 'mailtask.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Check current columns
cursor.execute("PRAGMA table_info(tasks)")
cols = [row[1] for row in cursor.fetchall()]
print(f"Current columns: {cols}")
print(f"Has email column: {'email' in cols}")

# Add email column if it doesn't exist
if 'email' not in cols:
    try:
        cursor.execute("ALTER TABLE tasks ADD COLUMN email TEXT")
        conn.commit()
        print("✓ Added email column successfully")
    except sqlite3.OperationalError as e:
        print(f"✗ Error adding column: {e}")
else:
    print("✓ Email column already exists")

# Verify
cursor.execute("PRAGMA table_info(tasks)")
cols = [row[1] for row in cursor.fetchall()]
print(f"Final columns: {cols}")
print(f"Has email column: {'email' in cols}")

conn.close()
print("\nDatabase fix complete!")

