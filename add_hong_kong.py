#!/usr/bin/env python
"""
Script to add Hong Kong to the countries table
"""

import sqlite3
from pathlib import Path

# Database path (same as in app.py)
DB_PATH = Path(__file__).resolve().parent / 'mailtask.db'

def add_hong_kong():
    """Add Hong Kong to countries table if it doesn't exist"""
    try:
        connection = sqlite3.connect(str(DB_PATH))
        connection.row_factory = sqlite3.Row
        cursor = connection.cursor()
        
        # Check if Hong Kong already exists
        cursor.execute("SELECT id, name FROM countries WHERE name = 'Hong Kong'")
        existing = cursor.fetchone()
        
        if existing:
            print(f"[INFO] Hong Kong already exists in database (ID: {existing['id']})")
        else:
            # Get max display_order
            cursor.execute("SELECT MAX(display_order) as max_order FROM countries")
            max_order = cursor.fetchone()['max_order'] or 0
            
            # Add Hong Kong
            cursor.execute("""
                INSERT INTO countries (name, display_order) VALUES (?, ?)
            """, ('Hong Kong', max_order + 1))
            connection.commit()
            print(f"[OK] Hong Kong added successfully with display_order: {max_order + 1}")
        
        # Verify
        cursor.execute("SELECT id, name, display_order FROM countries WHERE name = 'Hong Kong'")
        result = cursor.fetchone()
        if result:
            print(f"[OK] Verification:")
            print(f"  ID: {result['id']}")
            print(f"  Name: {result['name']}")
            print(f"  Display Order: {result['display_order']}")
        
        cursor.close()
        connection.close()
        return True
        
    except Exception as e:
        print(f"[ERROR] Error adding Hong Kong: {str(e)}")
        return False

if __name__ == '__main__':
    print("Adding Hong Kong to countries table...")
    print("-" * 50)
    success = add_hong_kong()
    
    if success:
        print("\n[SUCCESS] Hong Kong should now appear in the dropdown!")
        print("Please refresh the page to see the changes.")
    else:
        print("\n[FAILED] Could not add Hong Kong")


