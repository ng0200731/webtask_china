#!/usr/bin/env python
"""
Script to update user level in the database
Updates eric.brilliant@gmail.com to level 2
"""

import sqlite3
from pathlib import Path

# Database path (same as in app.py)
DB_PATH = Path(__file__).resolve().parent / 'mailtask.db'

def update_user_level(email: str, level: str, status: str = 'active'):
    """Update user level in the database"""
    if level not in ['1', '2', '3']:
        print(f"Error: Invalid level '{level}'. Must be 1, 2, or 3")
        return False
    
    if status not in ['active', 'suspended']:
        print(f"Error: Invalid status '{status}'. Must be active or suspended")
        return False
    
    try:
        connection = sqlite3.connect(str(DB_PATH))
        connection.row_factory = sqlite3.Row
        cursor = connection.cursor()
        
        # Check if user exists
        cursor.execute("SELECT id, level, status FROM users WHERE email = ?", (email,))
        user = cursor.fetchone()
        
        if user:
            # Update existing user
            old_level = user['level']
            old_status = user['status']
            cursor.execute("""
                UPDATE users 
                SET level = ?, status = ?
                WHERE email = ?
            """, (level, status, email))
            connection.commit()
            print(f"[OK] Updated user: {email}")
            print(f"  Level: {old_level} -> {level}")
            print(f"  Status: {old_status} -> {status}")
        else:
            # Create new user record
            from datetime import datetime
            created_at = datetime.utcnow().isoformat()
            cursor.execute("""
                INSERT INTO users (email, level, status, created_at, last_login, login_count)
                VALUES (?, ?, ?, ?, NULL, 0)
            """, (email, level, status, created_at))
            connection.commit()
            print(f"[OK] Created new user: {email}")
            print(f"  Level: {level}")
            print(f"  Status: {status}")
        
        # Verify the update
        cursor.execute("SELECT level, status FROM users WHERE email = ?", (email,))
        updated_user = cursor.fetchone()
        if updated_user:
            print(f"\n[OK] Verification:")
            print(f"  Email: {email}")
            print(f"  Level: {updated_user['level']}")
            print(f"  Status: {updated_user['status']}")
        
        cursor.close()
        connection.close()
        return True
        
    except Exception as e:
        print(f"[ERROR] Error updating user: {str(e)}")
        return False

if __name__ == '__main__':
    email = 'eric.brilliant@gmail.com'
    level = '2'
    status = 'active'
    
    print(f"Updating user level for: {email}")
    print(f"Setting level to: {level}")
    print(f"Setting status to: {status}")
    print("-" * 50)
    
    success = update_user_level(email, level, status)
    
    if success:
        print("\n[SUCCESS] Successfully updated user level!")
    else:
        print("\n[FAILED] Failed to update user level")

