#!/usr/bin/env python3
"""
Database initialization script for Railway deployment
Run this script to manually initialize the database tables
"""

import os
from dotenv import load_dotenv
from database import init_database

def main():
    """Initialize database tables"""
    load_dotenv()
    
    database_url = os.getenv('DATABASE_URL')
    if not database_url:
        print("❌ DATABASE_URL environment variable not found")
        print("Please set DATABASE_URL in your environment or .env file")
        return
    
    if database_url == 'postgresql://username:password@hostname:port/database':
        print("❌ DATABASE_URL is still set to placeholder value")
        print("Please update DATABASE_URL with your actual database connection string")
        return
    
    print(f"🔧 Initializing database...")
    print(f"📍 Database URL: {database_url[:30]}...")
    
    try:
        init_database()
        print("✅ Database initialization completed successfully!")
    except Exception as e:
        print(f"❌ Database initialization failed: {e}")
        return

if __name__ == "__main__":
    main()