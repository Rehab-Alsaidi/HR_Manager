import os
import psycopg2
from psycopg2.extras import DictCursor
from datetime import datetime
from typing import Optional, List, Dict, Any

def get_db_connection():
    """Get database connection using DATABASE_URL from environment"""
    database_url = os.getenv('DATABASE_URL')
    if not database_url:
        raise Exception("DATABASE_URL environment variable is required")
    
    try:
        conn = psycopg2.connect(database_url)
        return conn
    except Exception as e:
        print(f"Database connection failed: {e}")
        raise

def init_database():
    """Initialize database tables for HR evaluation system"""
    conn = get_db_connection()
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
                sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(employee_name, leader_email, evaluation_type, sent_date)
            )
        """)
        
        # Create employees table for caching employee data
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id SERIAL PRIMARY KEY,
                employee_name VARCHAR(255) NOT NULL,
                leader_name VARCHAR(255),
                leader_email VARCHAR(255),
                leader_crm VARCHAR(100),
                position VARCHAR(255),
                department VARCHAR(255),
                employee_crm VARCHAR(100),
                contract_renewal_date DATE,
                probation_end_date DATE,
                employee_status VARCHAR(100),
                last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(employee_name, leader_email)
            )
        """)
        
        # Create evaluation_reminders table for tracking sent reminders
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS evaluation_reminders (
                id SERIAL PRIMARY KEY,
                employee_name VARCHAR(255) NOT NULL,
                leader_name VARCHAR(255),
                leader_email VARCHAR(255) NOT NULL,
                evaluation_type VARCHAR(100) NOT NULL,
                deadline_date DATE NOT NULL,
                days_remaining INTEGER NOT NULL,
                email_sent BOOLEAN DEFAULT FALSE,
                email_sent_at TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # Create index for performance
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_sent_emails_date ON sent_emails(sent_date);
        """)
        
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_employees_email ON employees(leader_email);
        """)
        
        cursor.execute("""
            CREATE INDEX IF NOT EXISTS idx_reminders_date ON evaluation_reminders(deadline_date);
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
    cursor = conn.cursor()
    
    try:
        today = datetime.now().date()
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
    cursor = conn.cursor()
    
    try:
        today = datetime.now().date()
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
            WHERE sent_date < CURRENT_DATE - INTERVAL '%s days'
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

def sync_employee_data_to_db(employees_data: List[List[str]]):
    """Sync employee data from Feishu to database"""
    if not employees_data or len(employees_data) < 2:
        return
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        # Clear existing data
        cursor.execute("DELETE FROM employees")
        
        # Insert new data (skip header row)
        for row in employees_data[1:]:
            if len(row) < 10:
                continue
                
            employee_name = row[0] if row[0] else None
            leader_name = row[1] if row[1] else None
            contract_renewal = row[2] if row[2] else None
            probation_end = row[3] if row[3] else None
            employee_status = row[4] if row[4] else None
            position = row[5] if row[5] else None
            leader_email = row[6] if row[6] else None
            leader_crm = row[7] if row[7] else None
            department = row[8] if row[8] else None
            employee_crm = row[9] if row[9] else None
            
            if not employee_name or not leader_email:
                continue
            
            # Convert date strings to date objects
            contract_date = None
            probation_date = None
            
            try:
                if contract_renewal and contract_renewal != '':
                    contract_date = datetime.strptime(contract_renewal, '%Y-%m-%d').date()
            except:
                pass
                
            try:
                if probation_end and probation_end != '':
                    probation_date = datetime.strptime(probation_end, '%Y-%m-%d').date()
            except:
                pass
            
            cursor.execute("""
                INSERT INTO employees (
                    employee_name, leader_name, leader_email, leader_crm,
                    position, department, employee_crm, contract_renewal_date,
                    probation_end_date, employee_status
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (employee_name, leader_email) 
                DO UPDATE SET
                    leader_name = EXCLUDED.leader_name,
                    leader_crm = EXCLUDED.leader_crm,
                    position = EXCLUDED.position,
                    department = EXCLUDED.department,
                    employee_crm = EXCLUDED.employee_crm,
                    contract_renewal_date = EXCLUDED.contract_renewal_date,
                    probation_end_date = EXCLUDED.probation_end_date,
                    employee_status = EXCLUDED.employee_status,
                    last_updated = CURRENT_TIMESTAMP
            """, (
                employee_name, leader_name, leader_email, leader_crm,
                position, department, employee_crm, contract_date,
                probation_date, employee_status
            ))
        
        conn.commit()
        print(f"✅ Synced {len(employees_data)-1} employee records to database")
        
    except Exception as e:
        conn.rollback()
        print(f"❌ Error syncing employee data: {e}")
        raise
    finally:
        cursor.close()
        conn.close()

def get_employees_from_db() -> List[Dict[str, Any]]:
    """Get employee data from database"""
    conn = get_db_connection()
    cursor = conn.cursor(cursor_factory=DictCursor)
    
    try:
        cursor.execute("""
            SELECT * FROM employees 
            ORDER BY employee_name
        """)
        
        employees = []
        for row in cursor.fetchall():
            employees.append(dict(row))
        
        return employees
        
    except Exception as e:
        print(f"Error fetching employees from database: {e}")
        return []
    finally:
        cursor.close()
        conn.close()

def get_sent_emails_summary() -> Dict[str, Any]:
    """Get summary of sent emails from database"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        today = datetime.now().date()
        
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