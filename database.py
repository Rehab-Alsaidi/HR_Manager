import os
import psycopg2
from psycopg2.extras import DictCursor
from datetime import datetime, timezone
from typing import Optional, List, Dict, Any

def get_db_connection():
    """Get database connection using DATABASE_URL from environment"""
    database_url = os.getenv('DATABASE_URL')
    if not database_url or 'username:password@hostname' in database_url:
        # DATABASE_URL not configured or is placeholder - use file-based storage
        return None

    try:
        conn = psycopg2.connect(database_url)
        return conn
    except Exception as e:
        print(f"Database connection failed: {e}")
        return None

def init_database():
    """Initialize database tables for HR evaluation system"""
    conn = get_db_connection()
    if not conn:
        print("📝 Database not configured - using file-based storage")
        return

    cursor = conn.cursor()
    
    try:
        # Create sent_emails table to replace JSON file
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sent_emails (
                id SERIAL PRIMARY KEY,
                employee_name VARCHAR(255) NOT NULL,
                leader_email VARCHAR(255) NOT NULL,
                evaluation_type VARCHAR(100) NOT NULL,
                sent_date DATE NOT NULL,
                sent_at TIMESTAMPTZ DEFAULT (CURRENT_TIMESTAMP AT TIME ZONE 'UTC'),
                created_at TIMESTAMPTZ DEFAULT (CURRENT_TIMESTAMP AT TIME ZONE 'UTC'),
                UNIQUE(employee_name, leader_email, evaluation_type, sent_date)
            )
        """)
        
        # Create index for performance
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_sent_emails_date ON sent_emails(sent_date);
        """)
        
        conn.commit()
        print("✅ Database tables created successfully")
        
    except Exception as e:
        conn.rollback()
        print(f"❌ Error creating database tables: {e}")
        raise
    finally:
        cursor.close()
        conn.close()

def is_email_sent_today_db(employee_name: str, leader_email: str, evaluation_type: str) -> bool:
    """Check if email was already sent today using database"""
    conn = get_db_connection()
    if not conn:
        raise Exception("Database not available")

    cursor = conn.cursor()
    
    try:
        today = datetime.now(timezone.utc).date()
        cursor.execute("""
            SELECT COUNT(*) FROM sent_emails
            WHERE employee_name = %s
            AND leader_email = %s
            AND evaluation_type = %s
            AND sent_date = %s
        """, (employee_name, leader_email, evaluation_type, today))
        
        count = cursor.fetchone()[0]
        return count > 0
        
    except Exception as e:
        print(f"Error checking sent emails: {e}")
        return False
    finally:
        cursor.close()
        conn.close()

def mark_email_sent_db(employee_name: str, leader_email: str, evaluation_type: str):
    """Mark email as sent in database"""
    conn = get_db_connection()
    if not conn:
        raise Exception("Database not available")

    cursor = conn.cursor()
    
    try:
        today = datetime.now(timezone.utc).date()
        cursor.execute("""
            INSERT INTO sent_emails (employee_name, leader_email, evaluation_type, sent_date)
            VALUES (%s, %s, %s, %s)
            ON CONFLICT (employee_name, leader_email, evaluation_type, sent_date)
            DO NOTHING
        """, (employee_name, leader_email, evaluation_type, today))
        
        conn.commit()
        print(f"✅ Marked email as sent: {employee_name} -> {leader_email} ({evaluation_type})")
        
    except Exception as e:
        conn.rollback()
        print(f"❌ Error marking email as sent: {e}")
        raise
    finally:
        cursor.close()
        conn.close()

def cleanup_old_email_logs_db(days: int = 30):
    """Clean up old email logs from database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        cursor.execute("""
            DELETE FROM sent_emails
            WHERE sent_date < (CURRENT_DATE AT TIME ZONE 'UTC') - INTERVAL '%s days'
        """, (days,))
        
        deleted_count = cursor.rowcount
        conn.commit()
        
        if deleted_count > 0:
            print(f"✅ Cleaned up {deleted_count} old email logs")
            
    except Exception as e:
        conn.rollback()
        print(f"❌ Error cleaning up email logs: {e}")
    finally:
        cursor.close()
        conn.close()

# Removed sync_employee_data_to_db and get_employees_from_db functions
# App now reads directly from Base API, only uses database for email tracking

def get_sent_emails_summary() -> Dict[str, Any]:
    """Get summary of sent emails from database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        today = datetime.now(timezone.utc).date()

        # Get today's sent emails
        cursor.execute("""
            SELECT employee_name, leader_email, evaluation_type, sent_at
            FROM sent_emails
            WHERE sent_date = %s
            ORDER BY sent_at DESC
        """, (today,))
        
        today_emails = cursor.fetchall()
        
        # Get total count
        cursor.execute("SELECT COUNT(*) FROM sent_emails WHERE sent_date = %s", (today,))
        total_today = cursor.fetchone()[0]
        
        return {
            "today_date": today.isoformat(),
            "total_today": total_today,
            "today_emails": [
                {
                    "employee_name": row[0],
                    "leader_email": row[1],
                    "evaluation_type": row[2],
                    "sent_at": row[3].isoformat() if row[3] else None
                }
                for row in today_emails
            ]
        }
        
    except Exception as e:
        print(f"Error getting sent emails summary: {e}")
        return {"error": str(e)}
    finally:
        cursor.close()
        conn.close()