import json
import os
import random
import smtplib
import ssl
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any, Dict, List, Optional

import psycopg2
import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

from database import (
    cleanup_old_email_logs_db,
    get_sent_emails_summary,
    init_database,
    is_email_sent_today_db,
    mark_email_sent_db,
)

# Load environment variables
load_dotenv()

# Feishu API Base Configuration
VERBOSE = False
BASE = "https://open.feishu.cn/open-apis"

class FeishuError(Exception):
    pass

def set_verbose(value: bool) -> None:
    """Set verbose mode for debugging."""
    global VERBOSE
    VERBOSE = value

# Enable verbose mode for debugging
set_verbose(False)

def debug(message: str) -> None:
    """Print debug message if verbose mode is enabled."""
    if VERBOSE:
        print(message)

def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    """Get tenant access token from Feishu API."""
    url = f"{BASE}/auth/v3/tenant_access_token/internal"
    debug(f"[DEBUG] Getting tenant access token for app_id: {app_id}")
    
    payload = {"app_id": app_id, "app_secret": app_secret}
    r = requests.post(url, json=payload, timeout=20)
    r.raise_for_status()
    
    data = r.json()
    debug(f"[DEBUG] Token response: {data}")
    
    if data.get("code") != 0:
        error_msg = f"get_tenant_access_token failed: {data.get('code')} {data.get('msg')}"
        raise FeishuError(error_msg)
    
    debug("[DEBUG] Successfully got tenant access token")
    return data["tenant_access_token"]

def list_bitable_records(
    app_token: str,
    table_id: str,
    access_token: str,
    view_id: Optional[str] = None,
    page_token: Optional[str] = None,
    page_size: int = 500
) -> Dict[str, Any]:
    """List records from a Feishu bitable."""
    params: Dict[str, Any] = {'page_size': page_size}
    if view_id:
        params['view_id'] = view_id
    if page_token:
        params['page_token'] = page_token

    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    debug(f"[DEBUG] Listing bitable records: url={url} params={params}")
    
    try:
        headers = {"Authorization": f"Bearer {access_token}"}
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()
        
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable API error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable API error: {error_msg}")
        
        return resp.get('data', {})
    
    except requests.HTTPError as e:
        error_text = e.response.text
        print(f"[ERROR] HTTP error while listing bitable records: {error_text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {error_text}")
    
    except Exception as e:
        print(f"[ERROR] Unexpected error while listing bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_delete_bitable_records(
    app_token: str,
    table_id: str,
    record_ids: List[str],
    access_token: str
) -> Dict[str, Any]:
    """Delete multiple records from a Feishu bitable."""
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_delete"
    payload = {'record_ids': record_ids}
    debug(f"[DEBUG] Deleting bitable records: count={len(record_ids)}")
    
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable delete error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable delete error: {error_msg}")
        
        return resp
    
    except requests.HTTPError as e:
        error_text = e.response.text
        print(f"[ERROR] HTTP error while deleting bitable records: {error_text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {error_text}")
    
    except Exception as e:
        print(f"[ERROR] Unexpected error while deleting bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_create_bitable_records(
    app_token: str,
    table_id: str,
    records: List[Dict[str, Any]],
    access_token: str
) -> Dict[str, Any]:
    """Create multiple records in a Feishu bitable."""
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create"
    payload = {'records': records}
    debug(f"[DEBUG] Creating bitable records: count={len(records)}")
    
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable create error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable create error: {error_msg}")
        
        return resp
    
    except requests.HTTPError as e:
        error_text = e.response.text
        print(f"[ERROR] HTTP error while creating bitable records: {error_text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {error_text}")
    
    except Exception as e:
        print(f"[ERROR] Unexpected error while creating bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_update_bitable_records(
    app_token: str,
    table_id: str,
    records: List[Dict[str, Any]],
    access_token: str
) -> Dict[str, Any]:
    """Update multiple records in a Feishu bitable."""
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_update"
    payload = {'records': records}
    debug(f"[DEBUG] Updating bitable records: count={len(records)}")
    
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        r.raise_for_status()
        
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable update error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable update error: {error_msg}")
        
        return resp
    
    except requests.HTTPError as e:
        error_text = e.response.text
        print(f"[ERROR] HTTP error while updating bitable records: {error_text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {error_text}")
    
    except Exception as e:
        print(f"[ERROR] Unexpected error while updating bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

# Persistent duplicate prevention system
SENT_EMAILS_LOG = "sent_emails_log.json"

def load_sent_emails_log():
    """Load the log of previously sent emails."""
    try:
        if os.path.exists(SENT_EMAILS_LOG):
            with open(SENT_EMAILS_LOG, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading sent emails log: {e}")
    return {}

def save_sent_emails_log(log_data):
    """Save the log of sent emails."""
    try:
        with open(SENT_EMAILS_LOG, 'w') as f:
            json.dump(log_data, f, indent=2)
    except Exception as e:
        print(f"Error saving sent emails log: {e}")

def is_email_already_sent_today(employee_name, leader_email, evaluation_type):
    """Check if an email for this employee was already sent today (database or file)."""
    try:
        # Try database first
        database_url = os.getenv('DATABASE_URL')
        fallback_url = 'postgresql://username:password@hostname:port/database'
        if database_url and database_url != fallback_url:
            return is_email_sent_today_db(employee_name, leader_email, evaluation_type)
    except Exception as e:
        print(f"Database check failed, using file: {e}")
    
    # Fallback to file-based storage
    log_data = load_sent_emails_log()
    today = datetime.now().date().isoformat()
    
    # Create unique key for this employee-leader-evaluation combination
    key = f"{employee_name}|{leader_email}|{evaluation_type}"
    
    return log_data.get(today, {}).get(key, False)

def mark_email_as_sent(employee_name, leader_email, evaluation_type):
    """Mark an email as sent today (database or file)."""
    try:
        # Try database first
        database_url = os.getenv('DATABASE_URL')
        fallback_url = 'postgresql://username:password@hostname:port/database'
        if database_url and database_url != fallback_url:
            mark_email_sent_db(employee_name, leader_email, evaluation_type)
            return
    except Exception as e:
        print(f"Database marking failed, using file: {e}")
    
    # Fallback to file-based storage
    log_data = load_sent_emails_log()
    today = datetime.now().date().isoformat()
    
    if today not in log_data:
        log_data[today] = {}
    
    # Create unique key for this employee-leader-evaluation combination
    key = f"{employee_name}|{leader_email}|{evaluation_type}"
    log_data[today][key] = {
        "sent_at": datetime.now().isoformat(),
        "employee_name": employee_name,
        "leader_email": leader_email,
        "evaluation_type": evaluation_type
    }
    
    save_sent_emails_log(log_data)

def cleanup_old_logs():
    """Remove logs older than 30 days to prevent file from growing too large."""
    log_data = load_sent_emails_log()
    cutoff_date = (datetime.now().date() - timedelta(days=30)).isoformat()
    
    # Remove entries older than 30 days
    keys_to_remove = [date for date in log_data.keys() if date < cutoff_date]
    for key in keys_to_remove:
        del log_data[key]
    
    if keys_to_remove:
        save_sent_emails_log(log_data)

# Environment Variables for Lark Base Integration:
# LARK_APP_ID=your_app_id
# LARK_APP_SECRET=your_app_secret
# LARK_BASE_APP_TOKEN=your_base_app_token
# LARK_BASE_TABLE_ID=your_table_id
# LARK_BASE_VIEW_ID=your_view_id (optional)
# LARK_USE_USER_TOKEN=true/false (optional, defaults to false)
# LARK_USER_ACCESS_TOKEN=your_user_token (required if LARK_USE_USER_TOKEN=true)

app = Flask(__name__)

class LarkClient:
    def __init__(self):
        self.app_id = os.getenv('LARK_APP_ID')
        self.app_secret = os.getenv('LARK_APP_SECRET')
        
        if not self.app_id or not self.app_secret:
            raise Exception("LARK_APP_ID and LARK_APP_SECRET are required")
        
        # Store configuration for Base
        self.app_token = os.getenv('LARK_BASE_APP_TOKEN')
        self.table_id = os.getenv('LARK_BASE_TABLE_ID')
        self.view_id = os.getenv('LARK_BASE_VIEW_ID', '')
        self.use_user_token = os.getenv('LARK_USE_USER_TOKEN', 'false').lower() == 'true'
        
        # Initialize access token cache
        self.access_token = None
        self.token_expires = None
    
    def get_access_token(self):
        """Get tenant access token for Base API"""
        # If user token is configured, use it
        if self.use_user_token:
            user_token = os.getenv('LARK_USER_ACCESS_TOKEN')
            if user_token and user_token.strip() and user_token != '.':
                return user_token
            else:
                raise Exception("LARK_USER_ACCESS_TOKEN is required when LARK_USE_USER_TOKEN is true")
        
        # Always refresh tenant token to avoid invalid token errors
        # Use the new get_tenant_access_token function
        try:
            self.access_token = get_tenant_access_token(self.app_id, self.app_secret)
            self.token_expires = datetime.now() + timedelta(seconds=7200 - 300)  # 2 hours minus 5 minutes buffer
            return self.access_token
        except FeishuError as e:
            raise Exception(f"Failed to get access token: {e}")
    
    def get_base_data(self):
        """Get ALL data from Lark Base using the list_bitable_records function with pagination"""
        if not self.app_token or not self.table_id:
            raise Exception("LARK_BASE_APP_TOKEN and LARK_BASE_TABLE_ID are required for Base")
            
        access_token = self.get_access_token()
        
        # Convert Base record format to our expected format
        extracted_data = []
        
        # Add header row matching field_mappings order - KEEP COMPATIBILITY
        header_row = [
            "Employee Name", "Leader Name", "Contract Renewal Date", "Probation Period End Date", 
            "Employee Status", "Position", "Leader Email", "Leader CRM", "Department", "Employee CRM", 
            "Probation Remaining Days", "Contract Remaining Days", "Contract Company", "PSID", 
            "Big Team", "Small Team", "Marital Status", "Religion", "Joining Date", "2nd Contract Renewal",
            "Gender", "Nationality", "Birthday", "Age", "University", "Educational Level", 
            "School Ranking", "Major", "Exit Date", "Exit Type", "Exit Reason", "Work Email address", 
            "contract type", "service year", "Work Site"
        ]
        extracted_data.append(header_row)
        
        # Paginate through ALL records using the new function
        page_token = None
        total_records = 0
        
        while True:
            try:
                # Use the new list_bitable_records function
                response_data = list_bitable_records(
                    app_token=self.app_token,
                    table_id=self.table_id,
                    access_token=access_token,
                    view_id=self.view_id if self.view_id else None,
                    page_token=page_token,
                    page_size=500
                )
                
                # Extract records from response
                records = response_data.get('items', [])
                total_records += len(records)
                
                print(f"📥 Fetched {len(records)} records (Total: {total_records})")
                
                for record in records:
                    fields = record.get("fields", {})
                    
                    # Helper function to extract text from Base field values
                    def extract_base_field_value(field_value):
                        if isinstance(field_value, list) and len(field_value) > 0:
                            # Handle arrays like Position: ["CCSM"]
                            if isinstance(field_value[0], str):
                                return field_value[0]  # Return first item from array
                            # Handle rich text fields
                            elif isinstance(field_value[0], dict) and 'text' in field_value[0]:
                                return field_value[0]['text']
                            return str(field_value[0]) if field_value[0] else ""
                        elif isinstance(field_value, dict):
                            # Handle various field types
                            if 'text' in field_value:
                                return field_value['text']
                            elif 'name' in field_value:  # For select/multi-select fields
                                return field_value['name']
                        elif isinstance(field_value, str):
                            return field_value
                        elif isinstance(field_value, (int, float)):
                            return str(field_value)
                        elif field_value is not None:
                            return str(field_value)
                        return ""
                    
                    # Helper function to format dates from Base
                    def format_base_date(field_value):
                        # Handle timestamp directly if it's a number
                        if isinstance(field_value, (int, float)) and field_value > 1000000000:
                            try:
                                # Handle timestamp (milliseconds)
                                date_obj = datetime.fromtimestamp(field_value / 1000)
                                return date_obj.strftime('%Y-%m-%d')
                            except:
                                return ""
                        
                        # Handle string dates
                        date_str = extract_base_field_value(field_value)
                        if date_str and str(date_str).strip():
                            try:
                                date_str = str(date_str).strip()
                                # Base returns dates as timestamps or YYYY-MM-DD
                                if date_str.isdigit() and len(date_str) > 8:
                                    # Handle timestamp (milliseconds)
                                    date_obj = datetime.fromtimestamp(int(date_str) / 1000)
                                    return date_obj.strftime('%Y-%m-%d')
                                elif 'T' in date_str:
                                    # Handle ISO format
                                    date_obj = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                                    return date_obj.strftime('%Y-%m-%d')
                                else:
                                    # Already in YYYY-MM-DD format or other
                                    return date_str
                            except:
                                return date_str
                        return ""
                    
                    # Map actual Base column structure to expected output format
                    # Keep original structure for reminders compatibility 
                    field_mappings = [
                        ('Employee Name', ['Employee Name']),  # Position 0 - KEEP for reminders
                        ('Leader Name', ['Direct Leader CRM']),  # Position 1
                        ('Contract Renewal Date', ['1st Contract Renewal Date', 'Limited Contract End date']),  # Position 2
                        ('Probation Period End Date', ['Probation Period End Date']),  # Position 3
                        ('Employee Status', ['Employee Status']),  # Position 4
                        ('Position', ['Position']),  # Position 5
                        ('Leader Email', ['Direct Leader Email']),  # Position 6
                        ('Leader CRM', ['Direct Leader CRM']),  # Position 7
                        ('Department', ['Department']),  # Position 8
                        ('Employee CRM', ['CRM']),  # Position 9
                        ('Probation Remaining Days', ['Probation Period Remaining Days']),  # Position 10 - KEEP for reminders
                        ('Contract Remaining Days', ['Remaining Limited Contract End Days']),  # Position 11 - KEEP for reminders
                        # Add new fields for vendor notifications
                        ('Contract Company', ['Specific company name for signing the employment contract']),  # Position 12
                        ('PSID', ['PSID']),  # Position 13
                        ('Big Team', ['Big Team']),  # Position 14
                        ('Small Team', ['Small Team']),  # Position 15
                        ('Marital Status', ['Marital Status']),  # Position 16
                        ('Religion', ['Religion']),  # Position 17
                        ('Joining Date', ['Joining Date']),  # Position 18
                        ('2nd Contract Renewal', ['2nd Contract Renewal']),  # Position 19
                        ('Gender', ['Gender']),  # Position 20
                        ('Nationality', ['Nationality']),  # Position 21
                        ('Birthday', ['Birthday']),  # Position 22
                        ('Age', ['Age']),  # Position 23
                        ('University', ['University']),  # Position 24
                        ('Educational Level', ['Educational Level']),  # Position 25
                        ('School Ranking', ['School Ranking']),  # Position 26
                        ('Major', ['Major']),  # Position 27
                        ('Exit Date', ['Exit Date']),  # Position 28
                        ('Exit Type', ['Exit Type']),  # Position 29
                        ('Exit Reason', ['Exit Reason']),  # Position 30
                        ('Work Email address', ['Work Email address']),  # Position 31
                        ('contract type', ['contract type']),  # Position 32
                        ('service year', ['service year']),  # Position 33
                        ('Work Site', ['Work Site'])  # Position 34
                    ]
                    
                    extracted_row = []
                    for field_name, field_variations in field_mappings:
                        value = ""
                        # Try each field name variation
                        for field_var in field_variations:
                            if field_var in fields:
                                if field_name in ['Contract Renewal Date', 'Probation Period End Date']:
                                    value = format_base_date(fields[field_var])
                                else:
                                    value = extract_base_field_value(fields[field_var])
                                break
                        extracted_row.append(value)
                    
                    extracted_data.append(extracted_row)
                
                # Check if there are more pages
                has_more = response_data.get('has_more', False)
                page_token = response_data.get('page_token')
                
                if not has_more or not page_token:
                    break
                    
            except FeishuError as e:
                raise Exception(f"Failed to fetch Base data: {e}")
        
        print(f"✅ Base data loaded: {total_records} records total")
        return extracted_data

    
    def get_data(self):
        """Get data from Lark Base"""
        data = self.get_base_data()
        
        return data

def get_random_email_config():
    """Get a random email configuration from the available sender emails."""
    sender_emails_env = os.getenv('SENDER_EMAILS', os.getenv('SENDER_EMAIL'))
    sender_emails = [email.strip() for email in sender_emails_env.split(',')]

    usernames_env = os.getenv('EMAIL_USERNAMES', os.getenv('EMAIL_USERNAME'))
    email_usernames = [username.strip() for username in usernames_env.split(',')]

    passwords_env = os.getenv('EMAIL_PASSWORDS', os.getenv('EMAIL_PASSWORD'))
    email_passwords = [password.strip() for password in passwords_env.split(',')]

    # Use first email account to authenticate (usually the most reliable one)
    # But show all HR emails in the "From" field
    index = 0  # Use first account (sarakhateeb@51talk.com) for authentication

    return {
        'sender_email': ', '.join(sender_emails),  # Show all 3 emails in From field
        'auth_email': sender_emails[index],  # Email to use for SMTP authentication
        'username': email_usernames[index],
        'password': email_passwords[index]
    }

def send_reminder_email(
    employee_name,
    leader_name,
    leader_email,
    evaluation_link,
    days_remaining,
    evaluation_type,
    crm_account=""
):
    """Send evaluation reminder email to leader."""
    try:
        print(f"Attempting to send email to {leader_email} for {employee_name}")

        # Get email configuration (shows all 3 HR emails in From field)
        email_config = get_random_email_config()
        print(f"Using sender emails: {email_config['sender_email']}")
        
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', 465))
        
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
            server.login(email_config['username'], email_config['password'])

            message = MIMEMultipart()
            subject = f"Urgent: Employee Evaluation Required - {employee_name}"
            message["Subject"] = subject
            message["From"] = email_config['sender_email']  # Shows all 3 HR emails
            message["To"] = leader_email
            
            # Professional email body with HTML formatting for clickable links
            html_body = f"""
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #2c5aa0; border-bottom: 2px solid #2c5aa0; 
                   padding-bottom: 10px;">
            Employee Evaluation Required
        </h2>
        
        <p>Dear <strong>{leader_name}</strong>,</p>
        
        <p>This is an urgent reminder that <strong>{employee_name}'s</strong> 
           evaluation period is approaching and requires your immediate attention.</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; 
                    border-left: 4px solid #2c5aa0; margin: 20px 0;">
            <h3 style="margin-top: 0; color: #2c5aa0;">Employee Details:</h3>
            <ul style="margin: 10px 0;">
                <li><strong>Name:</strong> {employee_name}</li>
                <li><strong>CRM Account:</strong> {crm_account}</li>
                <li><strong>Evaluation Type:</strong> {evaluation_type}</li>
                <li><strong>Days Remaining:</strong> 
                    <span style="color: #dc3545; font-weight: bold;">
                        {days_remaining} days
                    </span>
                </li>
            </ul>
        </div>
        
        <div style="text-align: center; margin: 30px 0;">
            <a href="{evaluation_link}" 
               style="background-color: #2c5aa0; color: white; 
                      padding: 15px 30px; text-decoration: none; 
                      border-radius: 5px; font-weight: bold; 
                      display: inline-block; font-size: 16px;"
               target="_blank">
                📋 Complete Evaluation Form
            </a>
        </div>
        
        <p style="background-color: #fff3cd; padding: 10px; 
                  border: 1px solid #ffeaa7; border-radius: 4px;">
            <strong>⚠️ Important:</strong> Please complete this evaluation 
            before the deadline to ensure proper HR compliance.
        </p>
        
        <p>Thank you for your prompt attention to this matter.</p>
        
        <div style="margin-top: 30px; border-top: 1px solid #eee; 
                    padding-top: 15px; color: #666;">
            <p><strong>Best regards,</strong><br>
            HR Team<br>
            51Talk</p>
        </div>
    </div>
</body>
</html>
            """
            
            message.attach(MIMEText(html_body, "html"))
            server.send_message(message)
            print(f"Email sent successfully to {leader_email}")
            return True
            
    except Exception as e:
        print(f"Failed to send email to {leader_email}: {str(e)}")
        return False

def get_department_cc_emails(employees_data):
    """Get CC emails based on department mapping."""
    # Department-based CC mapping
    department_cc_mapping = {
        'CC': ['wuchuan@51talk.com'],
        'ACC': ['shichuan001@51talk.com'],
        'GCC': ['shichuan001@51talk.com'],
        'EA': ['guanshuhao001@51talk.com', 'nikiyang@51talk.com'],
        'CM': ['wangjingjing@51talk.com', 'nikiyang@51talk.com']
    }
    
    # Always include constant CC
    constant_cc = ['lijie14@51talk.com']
    
    # Collect all CC emails based on departments in this group
    cc_emails = set(constant_cc)  # Start with constant CC
    
    for emp in employees_data:
        department = emp.get('department', '').strip().upper()
        if department in department_cc_mapping:
            cc_emails.update(department_cc_mapping[department])
    
    return list(cc_emails)

def send_grouped_reminder_email(
    leader_name,
    leader_email,
    employees_data,
    evaluation_type,
    additional_cc_emails=None
):
    """Send one email to a leader with multiple employees listed."""
    try:
        employee_count = len(employees_data)
        print(f"Attempting to send grouped email to {leader_email} for {employee_count} employees")

        # Get email configuration (shows all 3 HR emails in From field)
        email_config = get_random_email_config()
        print(f"Using sender emails: {email_config['sender_email']}")
        
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', 465))
        
        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
            server.login(email_config['username'], email_config['password'])
            
            message = MIMEMultipart()
            message["Subject"] = "Urgent: Multiple Employee Evaluations Required"
            message["From"] = email_config['sender_email']
            message["To"] = leader_email
            
            # Get department-based CC emails
            department_cc_emails = get_department_cc_emails(employees_data)
            
            # Add any additional CC emails if provided
            all_cc_emails = department_cc_emails.copy()
            if additional_cc_emails:
                additional_emails = additional_cc_emails.split(',')
                additional_list = [
                    email.strip() for email in additional_emails 
                    if email.strip()
                ]
                all_cc_emails.extend(additional_list)
            
            # Remove duplicates and set CC header
            unique_cc_emails = list(set(all_cc_emails))
            if unique_cc_emails:
                message["Cc"] = ', '.join(unique_cc_emails)
            
            # Determine evaluation link and deadline
            if evaluation_type == "Probation Period Evaluation":
                evaluation_link = os.getenv('PROBATION_FORM_URL')
                email_intro = (
                    "Kindly find the link for the probation evaluation "
                    "to be done by your side before"
                )
                completion_note = (
                    "noting that if they pass this evaluation they will "
                    "be full-time employees"
                )
            else:
                evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                email_intro = (
                    "Kindly find below the name of your team employees that "
                    "need to be evaluated in order to renew their contracts "
                    "for a full year, in order to proceed timely it needs "
                    "to be done before"
                )
                completion_note = ""
            
            # Find the earliest deadline
            earliest_date = min(emp['deadline_date'] for emp in employees_data)
            
            # Create employee table rows
            employee_rows = ""
            for emp in employees_data:
                employee_rows += f"""
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('position', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;"><strong>{emp['name']}</strong></td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('employee_crm', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('department', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp['deadline_date']}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: #dc3545; font-weight: bold;">{emp['days_remaining']} days</td>
                </tr>"""
            
            # Professional email body with HTML formatting
            html_body = f"""
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 700px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #2c5aa0; border-bottom: 2px solid #2c5aa0; padding-bottom: 10px;">
            Employee Evaluations Required
        </h2>
        
        <p>Dear Leaders,</p>
        
        <p>{email_intro} <strong>{earliest_date}</strong>, for the following employees{'; ' + completion_note if completion_note else ''}:</p>
        
        <div style="text-align: center; margin: 20px 0;">
            <a href="{evaluation_link}" 
               style="background-color: #2c5aa0; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block; font-size: 16px;"
               target="_blank">
                📋 Complete Evaluation Form
            </a>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; border: 1px solid #ddd;">
            <thead>
                <tr style="background-color: #f8f9fa;">
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Position</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Employee Name</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">CRM</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Team</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Evaluation End Date</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Days Remaining</th>
                </tr>
            </thead>
            <tbody>
                {employee_rows}
            </tbody>
        </table>
        
        <div style="background-color: #fff3cd; padding: 15px; border: 1px solid #ffeaa7; border-radius: 4px; margin: 20px 0;">
            <p><strong>⚠️ Important:</strong> This evaluation must be completed within 2 working days. Please confirm by replying to this email once it's done.</p>
        </div>
        
        <p>Thank you for your prompt attention to this matter.</p>
        
        <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 15px; color: #666;">
            <p><strong>Best regards,</strong><br>
            HR Team<br>
            51Talk</p>
        </div>
    </div>
</body>
</html>
            """
            
            message.attach(MIMEText(html_body, "html"))
            
            # Get all recipients including CC
            recipients = [leader_email]
            if unique_cc_emails:
                recipients.extend(unique_cc_emails)

            # Use auth_email for SMTP envelope, but From header shows all 3 emails
            server.sendmail(email_config['auth_email'], recipients, message.as_string())
            print(f"Grouped email sent successfully to {leader_email} with CC: {', '.join(unique_cc_emails) if unique_cc_emails else 'None'}")
            return True
            
    except Exception as e:
        print(f"Failed to send grouped email to {leader_email}: {str(e)}")
        return False

def extract_email(email_data):
    """Extract email from various data formats - captures ALL data including '0' values"""
    if isinstance(email_data, list) and len(email_data) > 0:
        email_item = email_data[0]
        if isinstance(email_item, dict) and 'text' in email_item:
            return str(email_item['text']).strip() if email_item['text'] is not None else ""
        else:
            return str(email_item).strip() if email_item is not None else ""
    elif isinstance(email_data, str):
        return email_data.strip()
    elif email_data is not None:
        return str(email_data).strip()
    return ""

def is_valid_email_for_sending(email):
    """Check if email is valid for sending (not '0', empty, or invalid)"""
    if not email or email.strip() == "":
        return False
    email = email.strip()
    # Check for common invalid values
    invalid_values = ['0', 'null', 'none', 'n/a', '-', 'na']
    if email.lower() in invalid_values:
        return False
    # Basic email format check
    if '@' not in email or '.' not in email:
        return False
    return True

# Removed extract_evaluation_link function - now using static links

def excel_date_to_python(excel_date):
    if isinstance(excel_date, (int, float)):
        # Excel epoch is 1900-01-01, but Excel incorrectly treats 1900 as a leap year
        epoch = datetime(1899, 12, 30)
        return epoch + timedelta(days=excel_date)
    return None

def format_date_for_display(date_value):
    if isinstance(date_value, (int, float)):
        date_obj = excel_date_to_python(date_value)
        return date_obj.strftime('%Y-%m-%d') if date_obj else str(date_value)
    return str(date_value) if date_value else '-'

def check_and_send_reminders(employees_data, additional_cc_emails=None):
    sent_reminders = []
    today = datetime.now().date()
    
    # Group employees by leader email and evaluation type
    leader_groups = {}
    # Track processed employees to prevent duplicates
    processed_employees = set()
    
    for row_index, employee in enumerate(employees_data[1:], start=1):
        if len(employee) < 9:  # Now we need 9 columns for new structure
            continue
            
        employee_name = employee[0]  # Employee Name
        leader_name = employee[1]  # Leader Name (Direct Leader CRM)
        contract_renewal = employee[2]  # Contract Renewal Date
        probation_end = employee[3]  # Probation Period End Date
        employee_status = employee[4]  # Employee Status
        position = employee[5]  # Position
        leader_email_raw = employee[6]  # Leader Email
        leader_crm = employee[7]  # Leader CRM
        department = employee[8]  # Department
        employee_crm = employee[9]  # Employee CRM
        probation_remaining_days = employee[10] if len(employee) > 10 else None  # Probation Remaining Days
        contract_remaining_days = employee[11] if len(employee) > 11 else None  # Contract Remaining Days
        
        leader_email = extract_email(leader_email_raw)
        
        # Always capture employee data, but only process for email sending if we have valid data
        if not employee_name or not str(employee_name).strip():
            continue  # Skip if no employee name

        # Only process Active employees - skip all others (separated, terminated, etc.)
        if not employee_status or str(employee_status).strip().lower() != 'active':
            continue  # Skip non-active employees
        
        # Create unique key for this employee under this leader to prevent duplicates
        employee_leader_key = f"{employee_name.strip()}|{leader_email}"
        
        # Determine which evaluation is more urgent and closer to deadline
        chosen_evaluation = None
        chosen_date = None
        chosen_days = None
        
        # Check probation remaining days (direct from Base)
        if probation_remaining_days is not None:
            try:
                days_remaining = int(float(str(probation_remaining_days)))
                # Check for 1-20 days remaining
                if 1 <= days_remaining <= 20:
                    chosen_evaluation = "Probation Period Evaluation"
                    chosen_days = days_remaining
                    # Calculate the date
                    chosen_date = today + timedelta(days=days_remaining)
            except:
                pass
        
        # Check contract remaining days (direct from Base)
        if contract_remaining_days is not None:
            try:
                days_remaining = int(float(str(contract_remaining_days)))
                # Check for 1-20 days remaining
                if 1 <= days_remaining <= 20:
                    # If probation is also 20 days, probation takes priority
                    if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                        chosen_evaluation = "Contract Renewal Evaluation"
                        chosen_days = days_remaining
                        # Calculate the date
                        chosen_date = today + timedelta(days=days_remaining)
            except:
                pass
        
        # Only proceed if we have a valid evaluation (1-20 days) AND employee hasn't been processed yet
        if chosen_evaluation and employee_leader_key not in processed_employees:
            # Check if we have a valid email for sending
            if not is_valid_email_for_sending(leader_email):
                print(f"⚠️  Employee {employee_name} has invalid leader email '{leader_email}' - data captured but no email sent")
                # Still process the employee for data capture but don't send email
                processed_employees.add(employee_leader_key)
                continue
            
            # CRITICAL: Check if email was already sent today to prevent duplicates
            if is_email_already_sent_today(employee_name, leader_email, chosen_evaluation):
                print(f"SKIPPING {employee_name} - Email already sent today for {chosen_evaluation}")
                continue
                
            print(f"Processing employee: {employee_name} | Days remaining: {chosen_days} | Department: {department} | Type: {chosen_evaluation}")
            processed_employees.add(employee_leader_key)

            # Group by leader email + evaluation type + department (to keep departments separate)
            group_key = f"{leader_email}|{chosen_evaluation}|{department}"

            if group_key not in leader_groups:
                leader_groups[group_key] = {
                    'leader_name': leader_name or "Manager",
                    'leader_email': leader_email,
                    'evaluation_type': chosen_evaluation,
                    'department': department,
                    'employees': [],
                    'row_indices': [],
                    'employee_names': set()  # Track employee names in this group
                }
            
            # Double-check for duplicates within the same group
            if employee_name not in leader_groups[group_key]['employee_names']:
                leader_groups[group_key]['employee_names'].add(employee_name)
                leader_groups[group_key]['employees'].append({
                    'name': employee_name,
                    'position': position,  # Using actual position field
                    'department': department,
                    'employee_crm': employee_crm,
                    'deadline_date': chosen_date.strftime('%Y-%m-%d'),
                    'days_remaining': chosen_days,
                    'leader_crm': leader_crm
                })
                leader_groups[group_key]['row_indices'].append(row_index)

    # Send one email per group (leader + evaluation type + department)
    for group_key, group_data in leader_groups.items():
        email_sent = send_grouped_reminder_email(
            group_data['leader_name'],
            group_data['leader_email'], 
            group_data['employees'],
            group_data['evaluation_type'],
            additional_cc_emails
        )
        
        if email_sent:
            # CRITICAL: Mark all employees in this group as sent in persistent log
            for emp in group_data['employees']:
                mark_email_as_sent(emp['name'], group_data['leader_email'], group_data['evaluation_type'])
            
            # Update status for all employees in this group
            for row_index in group_data['row_indices']:
                employees_data[row_index][4] = "Email Sent"
            
            # Add to sent reminders
            for emp in group_data['employees']:
                sent_reminders.append({
                    "employee": emp['name'],
                    "leader": group_data['leader_name'],
                    "type": group_data['evaluation_type'].replace(' Evaluation', ''),
                    "days": emp['days_remaining'],
                    "email_sent": True
                })
    
    return sent_reminders

lark_client = LarkClient()

# Initialize database and cleanup old logs on startup
def initialize_app():
    """Initialize the application with database setup"""
    try:
        # Try to initialize database if DATABASE_URL is available
        database_url = os.getenv('DATABASE_URL')
        if database_url and database_url != 'postgresql://username:password@hostname:port/database':
            print("🔧 Initializing database...")
            init_database()
            cleanup_old_email_logs_db()
            print("✅ Database initialized successfully")
        else:
            print("📝 Using file-based storage (DATABASE_URL not configured)")
            cleanup_old_logs()
    except Exception as e:
        print(f"⚠️  Database initialization failed, using file-based storage: {e}")
        cleanup_old_logs()

# Initialize on startup
initialize_app()


@app.route('/vendor-notifications')
def vendor_notifications():
    """Vendor notifications page"""
    return render_template('vendor_notifications.html')

@app.route('/reminders')
def todays_reminders():
    try:
        data = lark_client.get_data()
        today = datetime.now().date()
        
        # Group employees by leader email and evaluation type
        # Include both: pending reminders AND emails sent today
        leader_groups = {}
        processed_employees = set()
        
        for row_index, employee in enumerate(data[1:], start=1):
            if len(employee) < 12:  # Need 12 columns for remaining days fields
                continue
                
            employee_name = employee[0]
            leader_name = employee[1]
            contract_renewal = employee[2]
            probation_end = employee[3]
            employee_status = employee[4]
            position = employee[5]
            leader_email_raw = employee[6]
            leader_crm = employee[7]
            department = employee[8]
            employee_crm = employee[9]
            
            leader_email = extract_email(leader_email_raw)
            
            if not employee_name or not str(employee_name).strip():
                continue  # Skip if no employee name

            # Only show Active employees in reminders view - skip all others
            if not employee_status or str(employee_status).strip().lower() != 'active':
                continue  # Skip non-active employees
            
            employee_leader_key = f"{employee_name.strip()}|{leader_email}"
            
            chosen_evaluation = None
            chosen_date = None
            chosen_days = None
            
            # Get remaining days fields (same logic as check_and_send_reminders)
            probation_remaining_days = employee[10] if len(employee) > 10 else None
            contract_remaining_days = employee[11] if len(employee) > 11 else None
            
            # Check probation remaining days (direct from Base)
            if probation_remaining_days is not None:
                try:
                    days_remaining = int(float(str(probation_remaining_days)))
                    # Check for 1-20 days remaining
                    if 1 <= days_remaining <= 20:
                        chosen_evaluation = "Probation Period Evaluation"
                        chosen_days = days_remaining
                        # Calculate the date
                        chosen_date = today + timedelta(days=days_remaining)
                except:
                    pass

            # Check contract remaining days (direct from Base)
            if contract_remaining_days is not None:
                try:
                    days_remaining = int(float(str(contract_remaining_days)))
                    # Check for 1-20 days remaining
                    if 1 <= days_remaining <= 20:
                        # If probation is also in range, probation takes priority
                        if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_days = days_remaining
                            # Calculate the date
                            chosen_date = today + timedelta(days=days_remaining)
                except:
                    pass

            # Also check for employees with "Email Sent" status (recently processed)
            email_sent_recently = (employee_status == "Email Sent")

            # Include if: 1) Valid evaluation in 1-20 days, OR 2) Email was sent recently (today)
            should_include = (chosen_evaluation and employee_leader_key not in processed_employees) or email_sent_recently
            
            if should_include:
                processed_employees.add(employee_leader_key)
                
                # For email sent cases, we need to reconstruct the evaluation type and date
                if email_sent_recently and not chosen_evaluation:
                    # Try to determine what type of evaluation was sent by checking dates
                    if probation_end:
                        try:
                            if isinstance(probation_end, (int, float)):
                                eval_date = excel_date_to_python(probation_end).date()
                            else:
                                eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                            chosen_evaluation = "Probation Period Evaluation"
                            chosen_date = eval_date
                            chosen_days = (eval_date - today).days
                        except:
                            pass
                    
                    if not chosen_evaluation and contract_renewal:
                        try:
                            if isinstance(contract_renewal, (int, float)):
                                eval_date = excel_date_to_python(contract_renewal).date()
                            else:
                                eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_date = eval_date
                            chosen_days = (eval_date - today).days
                        except:
                            pass
                
                # Group by leader email + evaluation type + department (to keep departments separate)
                group_key = f"{leader_email}|{chosen_evaluation}|{department}"

                if group_key not in leader_groups:
                    leader_groups[group_key] = {
                        'leader_name': leader_name or "Manager",
                        'leader_email': leader_email,
                        'evaluation_type': chosen_evaluation,
                        'department': department,
                        'employees': [],
                        'employee_names': set()
                    }

                if employee_name not in leader_groups[group_key]['employee_names']:
                    leader_groups[group_key]['employee_names'].add(employee_name)
                    leader_groups[group_key]['employees'].append({
                        'name': employee_name,
                        'department': department,
                        'deadline_date': chosen_date.strftime('%Y-%m-%d') if chosen_date else 'N/A',
                        'days_remaining': chosen_days if chosen_days is not None else 'N/A',
                        'position': position,
                        'employee_crm': employee_crm,
                        'leader_crm': leader_crm,
                        'email_sent': email_sent_recently  # Add flag to indicate if email was already sent
                    })
        
        # Convert to list for template
        reminders_data = []
        for group_key, group_data in leader_groups.items():
            # Get department-based CC emails for this group
            department_cc_emails = get_department_cc_emails(group_data['employees'])

            reminders_data.append({
                'leader_name': group_data['leader_name'],
                'leader_email': group_data['leader_email'],
                'evaluation_type': group_data['evaluation_type'],
                'employee_count': len(group_data['employees']),
                'employees': group_data['employees'],
                'department': group_data['department'],  # Single department now
                'cc_emails': ', '.join(department_cc_emails)
            })
        
        # Sort by urgency (earliest deadline first)
        reminders_data.sort(key=lambda x: min(emp['days_remaining'] for emp in x['employees']))
        
        return render_template('reminders.html', 
                             reminders=reminders_data, 
                             total_reminders=len(reminders_data),
                             total_employees=sum(r['employee_count'] for r in reminders_data),
                             today_date=today.strftime('%Y-%m-%d'))
    except Exception as e:
        return f"Error loading reminders: {str(e)}", 500

@app.route('/')
def index():
    try:
        data = lark_client.get_data()
        # Process data to format dates for display
        formatted_data = []
        total_employees = 0
        
        if data:
            # Keep header row as is
            formatted_data.append(data[0])
            
            # Format employee data and count actual employees
            for employee in data[1:]:
                if len(employee) >= 9:
                    # Check if this is a real employee record (has employee name and it's not empty)
                    employee_name = str(employee[0]).strip() if employee[0] else ""
                    if employee_name and employee_name not in ['-', 'null', 'None', '']:  # Employee name exists and is meaningful
                        total_employees += 1
                        
                        # Determine which evaluation link to show based on which date is closer
                        probation_date = employee[3] if employee[3] else None
                        contract_date = employee[2] if employee[2] else None
                        evaluation_link = os.getenv('PROBATION_FORM_URL')  # Default to probation form
                        
                        # Choose appropriate link based on upcoming dates
                        if probation_date and contract_date:
                            try:
                                today = datetime.now().date()
                                prob_date = excel_date_to_python(probation_date).date() if isinstance(probation_date, (int, float)) else datetime.strptime(str(probation_date), "%Y-%m-%d").date()
                                cont_date = excel_date_to_python(contract_date).date() if isinstance(contract_date, (int, float)) else datetime.strptime(str(contract_date), "%Y-%m-%d").date()
                                
                                prob_days = (prob_date - today).days
                                cont_days = (cont_date - today).days
                                
                                # Use contract renewal link if it's closer and within 20 days
                                if 0 <= cont_days <= 20 and (prob_days > 20 or cont_days < prob_days):
                                    evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                            except:
                                pass  # Keep default probation link
                        elif contract_date:
                            evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                        
                        formatted_employee = [
                            employee[0],  # Employee Name (Column I)
                            employee[1] or "Manager",  # Leader Name (empty in new structure)
                            format_date_for_display(employee[3]),  # Probation End Date (Column P)
                            format_date_for_display(employee[2]),  # Contract Renewal Date (Column O)
                            evaluation_link,  # Dynamic Evaluation Link
                            employee[4],  # Employee Status (Column AB)
                            employee[5],  # Position (Column L)
                            employee[6],  # Leader Email (Column G)
                            employee[7],  # Leader CRM (Column F)
                            employee[8],  # Department (Column H)
                            employee[9]   # Employee CRM (Column C)
                        ]
                        formatted_data.append(formatted_employee)
        
        return render_template('index.html', employees=formatted_data, total_employees=total_employees)
    except Exception as e:
        return f"Error fetching data: {str(e)}", 500

@app.route('/api/employees')
def api_employees():
    try:
        data = lark_client.get_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/send-reminders', methods=['POST'])
def send_reminders():
    try:
        # Get CC emails from request
        request_data = request.get_json() if request.is_json else {}
        additional_cc_emails = request_data.get('cc_emails', '')
        
        data = lark_client.get_data()
        sent = check_and_send_reminders(data, additional_cc_emails)
        return jsonify({"success": True, "sent_reminders": sent})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/debug-data')
def debug_data():
    """Debug endpoint to see raw data structure"""
    try:
        data = lark_client.get_data()
        return jsonify({
            "success": True,
            "total_rows": len(data),
            "header": data[0] if len(data) > 0 else [],
            "first_few_rows": data[1:4] if len(data) > 1 else [],
            "column_count": len(data[0]) if len(data) > 0 else 0
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/debug-sent-emails')
def debug_sent_emails():
    """Debug endpoint to see sent emails log for duplicate prevention"""
    try:
        # Try database first
        database_url = os.getenv('DATABASE_URL')
        if database_url and database_url != 'postgresql://username:password@hostname:port/database':
            return jsonify({
                "success": True,
                "source": "database",
                **get_sent_emails_summary()
            })
        
        # Fallback to file
        log_data = load_sent_emails_log()
        today = datetime.now().date().isoformat()
        
        return jsonify({
            "success": True,
            "source": "file",
            "today": today,
            "todays_sent_emails": log_data.get(today, {}),
            "all_dates": list(log_data.keys()),
            "total_entries_today": len(log_data.get(today, {}))
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/test-duplicate-prevention', methods=['POST'])
def test_duplicate_prevention():
    """Test endpoint to verify duplicate prevention works"""
    try:
        # Test the duplicate prevention logic without actually sending emails
        data = lark_client.get_data()
        today = datetime.now().date()
        
        test_results = []
        processed_count = 0
        skipped_count = 0
        
        for employee in data[1:6]:  # Test first 5 employees only
            if len(employee) < 10:
                continue
                
            employee_name = employee[0]
            leader_email_raw = employee[6]
            leader_email = extract_email(leader_email_raw)
            
            if not leader_email or not employee_name:
                continue
            
            # Check probation end date for test
            probation_end = employee[3]
            if probation_end:
                try:
                    if isinstance(probation_end, (int, float)):
                        eval_date = excel_date_to_python(probation_end).date()
                    else:
                        eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                    
                    days_until = (eval_date - today).days

                    if 1 <= days_until <= 20:
                        evaluation_type = "Probation Period Evaluation"
                        
                        # Check if would be skipped due to duplicate
                        would_skip = is_email_already_sent_today(employee_name, leader_email, evaluation_type)
                        
                        test_results.append({
                            "employee": employee_name,
                            "leader_email": leader_email,
                            "evaluation_type": evaluation_type,
                            "days_remaining": days_until,
                            "would_skip_duplicate": would_skip,
                            "status": "SKIP" if would_skip else "SEND"
                        })
                        
                        if would_skip:
                            skipped_count += 1
                        else:
                            processed_count += 1
                except:
                    pass
        
        return jsonify({
            "success": True,
            "test_results": test_results,
            "summary": {
                "would_send": processed_count,
                "would_skip": skipped_count,
                "total_tested": len(test_results)
            }
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

@app.route('/api/preview-reminders')
def preview_reminders():
    """Preview which emails would be sent without actually sending them"""
    try:
        data = lark_client.get_data()
# Debug output removed - system working correctly
        today = datetime.now().date()
        
        # Group employees by leader email, evaluation type, and department (EXACT same logic as check_and_send_reminders)
        leader_groups = {}
        # Track processed employees to prevent duplicates
        processed_employees = set()

        for row_index, employee in enumerate(data[1:], start=1):
            if len(employee) < 12:  # Need 12 columns for remaining days fields
                continue

            employee_name = employee[0]
            leader_name = employee[1]
            contract_renewal = employee[2]
            probation_end = employee[3]
            employee_status = employee[4]
            position = employee[5]
            leader_email_raw = employee[6]
            leader_crm = employee[7]
            department = employee[8]
            employee_crm = employee[9]
            probation_remaining_days = employee[10] if len(employee) > 10 else None
            contract_remaining_days = employee[11] if len(employee) > 11 else None

            leader_email = extract_email(leader_email_raw)

            if not employee_name or not str(employee_name).strip():
                continue  # Skip if no employee name

            # Only preview Active employees - skip all others (separated, terminated, etc.)
            if not employee_status or str(employee_status).strip().lower() != 'active':
                continue  # Skip non-active employees

            # Create unique key for this employee under this leader to prevent duplicates
            employee_leader_key = f"{employee_name.strip()}|{leader_email}"

            # Determine which evaluation is more urgent and closer to deadline
            chosen_evaluation = None
            chosen_date = None
            chosen_days = None

            # Check probation remaining days (direct from Base)
            if probation_remaining_days is not None:
                try:
                    days_remaining = int(float(str(probation_remaining_days)))
                    # Check for 1-20 days remaining
                    if 1 <= days_remaining <= 20:
                        chosen_evaluation = "Probation Period Evaluation"
                        chosen_days = days_remaining
                        # Calculate the date
                        chosen_date = today + timedelta(days=days_remaining)
                except:
                    pass

            # Check contract remaining days (direct from Base)
            if contract_remaining_days is not None:
                try:
                    days_remaining = int(float(str(contract_remaining_days)))
                    # Check for 1-20 days remaining
                    if 1 <= days_remaining <= 20:
                        # If probation is also in range, probation takes priority
                        if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_days = days_remaining
                            # Calculate the date
                            chosen_date = today + timedelta(days=days_remaining)
                except:
                    pass

            # Only proceed if we have a valid evaluation (1-20 days) AND employee hasn't been processed yet
            if chosen_evaluation and employee_leader_key not in processed_employees:
                # Check if we have a valid email for sending
                if not is_valid_email_for_sending(leader_email):
                    print(f"PREVIEW SKIP: {employee_name} has invalid leader email '{leader_email}'")
                    continue

                # CRITICAL: Check if email was already sent today to prevent duplicates in preview
                if is_email_already_sent_today(employee_name, leader_email, chosen_evaluation):
                    print(f"PREVIEW SKIP: {employee_name} - Email already sent today for {chosen_evaluation}")
                    continue

                print(f"Preview: {employee_name} | Days remaining: {chosen_days} | Department: {department} | Type: {chosen_evaluation}")
                processed_employees.add(employee_leader_key)

                # Group by leader email + evaluation type + department (to keep departments separate)
                group_key = f"{leader_email}|{chosen_evaluation}|{department}"

                if group_key not in leader_groups:
                    leader_groups[group_key] = {
                        'leader_name': leader_name or "Manager",
                        'leader_email': leader_email,
                        'evaluation_type': chosen_evaluation,
                        'department': department,
                        'employees': [],
                        'employee_names': set()  # Track employee names in this group
                    }

                # Double-check for duplicates within the same group
                if employee_name not in leader_groups[group_key]['employee_names']:
                    leader_groups[group_key]['employee_names'].add(employee_name)
                    leader_groups[group_key]['employees'].append({
                        'name': employee_name,
                        'department': department,
                        'deadline_date': chosen_date.strftime('%Y-%m-%d'),
                        'days_remaining': chosen_days
                    })
        
        preview_data = []
        for group_key, group_data in leader_groups.items():
            preview_data.append({
                'leader_name': group_data['leader_name'],
                'leader_email': group_data['leader_email'],
                'evaluation_type': group_data['evaluation_type'],
                'department': group_data['department'],
                'employee_count': len(group_data['employees']),
                'employees': group_data['employees']
            })
        
        return jsonify({"success": True, "preview": preview_data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/database-status')
def database_status():
    """Check database connection and status"""
    try:
        database_url = os.getenv('DATABASE_URL')
        if not database_url or database_url == 'postgresql://username:password@hostname:port/database':
            return jsonify({
                "success": True,
                "status": "disabled",
                "message": "Database not configured, using file-based storage"
            })
        
        # Test database connection
        from database import get_db_connection
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT 1")
        cursor.close()
        conn.close()
        
        return jsonify({
            "success": True,
            "status": "connected",
            "message": "Database connection successful",
            "database_url": database_url[:30] + "..." if len(database_url) > 30 else database_url
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "status": "error",
            "error": str(e)
        })

@app.route('/api/database-verify-tables')
def verify_database_tables():
    """Verify all database tables exist and show their structure"""
    try:
        database_url = os.getenv('DATABASE_URL')
        if not database_url or database_url == 'postgresql://username:password@hostname:port/database':
            return jsonify({
                "success": False,
                "message": "Database not configured"
            })
        
        from database import get_db_connection
        conn = get_db_connection()
        cursor = conn.cursor()
        
        # Check if tables exist
        tables_to_check = ['sent_emails', 'employees', 'evaluation_reminders']
        table_info = {}
        
        for table_name in tables_to_check:
            # Check if table exists
            cursor.execute("""
                SELECT COUNT(*) 
                FROM information_schema.tables 
                WHERE table_name = %s
            """, (table_name,))
            
            exists = cursor.fetchone()[0] > 0
            
            if exists:
                # Get table structure
                cursor.execute("""
                    SELECT column_name, data_type, is_nullable, column_default
                    FROM information_schema.columns 
                    WHERE table_name = %s
                    ORDER BY ordinal_position
                """, (table_name,))
                
                columns = cursor.fetchall()
                
                # Get row count
                cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                row_count = cursor.fetchone()[0]
                
                table_info[table_name] = {
                    "exists": True,
                    "row_count": row_count,
                    "columns": [
                        {
                            "name": col[0],
                            "type": col[1],
                            "nullable": col[2],
                            "default": col[3]
                        }
                        for col in columns
                    ]
                }
            else:
                table_info[table_name] = {
                    "exists": False,
                    "row_count": 0,
                    "columns": []
                }
        
        cursor.close()
        conn.close()
        
        # Check if all required tables exist
        all_tables_exist = all(info["exists"] for info in table_info.values())
        
        return jsonify({
            "success": True,
            "all_tables_exist": all_tables_exist,
            "tables": table_info,
            "database_url": database_url[:50] + "..." if len(database_url) > 50 else database_url
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/database-init', methods=['POST'])
def init_database_endpoint():
    """Manually initialize database tables"""
    try:
        database_url = os.getenv('DATABASE_URL')
        if not database_url or database_url == 'postgresql://username:password@hostname:port/database':
            return jsonify({
                "success": False,
                "message": "Database not configured"
            })
        
        from database import init_database
        init_database()
        
        return jsonify({
            "success": True,
            "message": "Database tables initialized successfully"
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/database-sync-test', methods=['POST'])
def test_database_sync():
    """Test syncing current Feishu data to database"""
    try:
        database_url = os.getenv('DATABASE_URL')
        if not database_url or database_url == 'postgresql://username:password@hostname:port/database':
            return jsonify({
                "success": False,
                "message": "Database not configured"
            })
        
        # Get current data from Feishu
        data = lark_client.get_data()
        
        # Count employees directly from Base data
        employee_count = len(data) - 1 if len(data) > 1 else 0
        
        return jsonify({
            "success": True,
            "message": "Data synced successfully",
            "employees_synced": employee_count,
            "total_rows_from_feishu": len(data) - 1  # Subtract header row
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/debug-railway-issue')
def debug_railway_issue():
    """Debug why Railway shows different results than local"""
    try:
        from datetime import datetime
        
        # Check environment
        env_check = {
            "LARK_APP_ID": os.getenv('LARK_APP_ID', 'NOT_SET'),
            "LARK_APP_SECRET": os.getenv('LARK_APP_SECRET', 'NOT_SET')[:10] + "..." if os.getenv('LARK_APP_SECRET') else 'NOT_SET',
            "LARK_BASE_APP_TOKEN": os.getenv('LARK_BASE_APP_TOKEN', 'NOT_SET'),
            "LARK_BASE_TABLE_ID": os.getenv('LARK_BASE_TABLE_ID', 'NOT_SET'),
            "DATABASE_URL": "SET" if os.getenv('DATABASE_URL') and os.getenv('DATABASE_URL') != 'postgresql://username:password@hostname:port/database' else 'NOT_SET'
        }
        
        # Check current date/time
        now = datetime.now()
        today = now.date()
        
        # Try to get data
        try:
            data = lark_client.get_data()
            data_source = "SUCCESS"
            total_rows = len(data) if data else 0
            has_employees = total_rows > 1
        except Exception as e:
            data_source = f"ERROR: {str(e)}"
            total_rows = 0
            has_employees = False
            data = []
        
        # Check for employees with 20 days remaining
        employees_20_days = []
        if has_employees:
            for employee in data[1:]:  # Skip header
                if len(employee) < 10:
                    continue
                    
                employee_name = employee[0]
                if not employee_name:
                    continue
                    
                probation_end = employee[3]
                contract_renewal = employee[2]
                leader_email = extract_email(employee[6]) if len(employee) > 6 else ""
                
                # Check dates
                for date_field, eval_type in [(probation_end, "Probation"), (contract_renewal, "Contract")]:
                    if date_field:
                        try:
                            if isinstance(date_field, (int, float)):
                                eval_date = excel_date_to_python(date_field).date()
                            else:
                                eval_date = datetime.strptime(str(date_field), "%Y-%m-%d").date()
                            
                            days_until = (eval_date - today).days
                            
                            if days_until == 20:
                                employees_20_days.append({
                                    "name": employee_name,
                                    "eval_type": eval_type,
                                    "eval_date": eval_date.isoformat(),
                                    "days_until": days_until,
                                    "leader_email": leader_email,
                                    "email_valid": is_valid_email_for_sending(leader_email)
                                })
                        except:
                            pass
        
        return jsonify({
            "success": True,
            "server_time": now.isoformat(),
            "server_date": today.isoformat(),
            "environment_variables": env_check,
            "data_source_status": data_source,
            "total_data_rows": total_rows,
            "employees_with_20_days": employees_20_days,
            "should_send_reminders": len(employees_20_days) > 0
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        })

@app.route('/api/debug-data-source')
def debug_data_source():
    """Debug endpoint to check which data source is being used"""
    try:
        
        # Check Base configuration
        base_config = {
            'app_token': os.getenv('LARK_BASE_APP_TOKEN'),
            'table_id': os.getenv('LARK_BASE_TABLE_ID'),
            'view_id': os.getenv('LARK_BASE_VIEW_ID', ''),
            'use_user_token': os.getenv('LARK_USE_USER_TOKEN', 'false').lower() == 'true'
        }
        
        # Try Base API
        base_result = None
        base_error = None
        try:
            data = lark_client.get_base_data()
            base_result = f"SUCCESS - {len(data)} rows"
        except Exception as e:
            base_error = str(e)
            base_result = f"FAILED - {base_error}"
        
        
        # What get_data() actually returns
        actual_data = lark_client.get_data()
        actual_result = f"{len(actual_data)} rows"
        
        return jsonify({
            "success": True,
            "base_config": base_config,
            "base_api_test": base_result,
            "base_error": base_error,
            "actual_data_returned": actual_result,
            "environment": {
                "app_id": os.getenv('LARK_APP_ID', '')[:10] + "..." if os.getenv('LARK_APP_ID') else None,
                "app_secret": "SET" if os.getenv('LARK_APP_SECRET') else "NOT SET"
            }
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/debug-railway-complete')
def debug_railway_complete():
    """Complete debug information for Railway deployment"""
    try:
        today = datetime.now().date()
        debug_info = {
            "timestamp": datetime.now().isoformat(),
            "today": str(today),
            "step1_data_fetch": {},
            "step2_column_check": {},
            "step3_evaluation_logic": {},
            "step4_email_check": {}
        }
        
        # Step 1: Data Fetch
        try:
            data = lark_client.get_data()
            debug_info["step1_data_fetch"] = {
                "success": True,
                "total_rows": len(data),
                "has_header": len(data) > 0,
                "headers": data[0] if len(data) > 0 else [],
                "sample_row_length": len(data[1]) if len(data) > 1 else 0
            }
        except Exception as e:
            debug_info["step1_data_fetch"] = {"success": False, "error": str(e)}
            return jsonify(debug_info)
        
        # Step 2: Column Check
        if len(data) > 1:
            sample_row = data[1]
            debug_info["step2_column_check"] = {
                "row_length": len(sample_row),
                "has_12_columns": len(sample_row) >= 12,
                "probation_days_column": sample_row[10] if len(sample_row) > 10 else "MISSING",
                "contract_days_column": sample_row[11] if len(sample_row) > 11 else "MISSING",
                "employee_name": sample_row[0] if len(sample_row) > 0 else "MISSING"
            }
        
        # Step 3: Evaluation Logic
        employees_checked = 0
        employees_with_days = 0
        employees_in_range = 0
        sample_found = []
        
        for employee in data[1:]:
            if len(employee) < 12:
                continue
            employees_checked += 1
            
            prob_days = employee[10] if employee[10] and employee[10] != '' else None
            cont_days = employee[11] if employee[11] and employee[11] != '' else None
            
            has_days = False
            if prob_days or cont_days:
                employees_with_days += 1
                has_days = True
            
            if prob_days:
                try:
                    days = int(float(str(prob_days)))
                    if 1 <= days <= 20:
                        employees_in_range += 1
                        if len(sample_found) < 3:
                            sample_found.append({
                                "name": employee[0],
                                "type": "Probation",
                                "days": days,
                                "leader_email": employee[6] if len(employee) > 6 else "MISSING"
                            })
                except:
                    pass

            if cont_days:
                try:
                    days = int(float(str(cont_days)))
                    if 1 <= days <= 20:
                        employees_in_range += 1
                        if len(sample_found) < 3:
                            sample_found.append({
                                "name": employee[0],
                                "type": "Contract",
                                "days": days,
                                "leader_email": employee[6] if len(employee) > 6 else "MISSING"
                            })
                except:
                    pass

        debug_info["step3_evaluation_logic"] = {
            "employees_checked": employees_checked,
            "employees_with_days_data": employees_with_days,
            "employees_in_range_1_20": employees_in_range,
            "sample_found": sample_found
        }
        
        # Step 4: Email Check
        try:
            from app import check_and_send_reminders
            
            # Mock email sending to count
            original_send = globals().get('send_grouped_reminder_email')
            email_count = 0
            def mock_send(*args, **kwargs):
                nonlocal email_count
                email_count += 1
                return True
            
            if 'send_grouped_reminder_email' in globals():
                globals()['send_grouped_reminder_email'] = mock_send
            
            reminders = check_and_send_reminders(data)
            
            debug_info["step4_email_check"] = {
                "reminders_generated": len(reminders),
                "emails_would_send": email_count,
                "sample_reminders": reminders[:3] if reminders else []
            }
            
            # Restore original
            if original_send:
                globals()['send_grouped_reminder_email'] = original_send
                
        except Exception as e:
            debug_info["step4_email_check"] = {"error": str(e)}
        
        return jsonify(debug_info)
        
    except Exception as e:
        return jsonify({"error": str(e), "timestamp": datetime.now().isoformat()})

@app.route('/api/debug-evaluation-logic')
def debug_evaluation_logic():
    """Debug the evaluation logic specifically"""
    try:
        data = lark_client.get_data()
        today = datetime.now().date()
        
        debug_info = {
            "total_rows": len(data),
            "today": str(today),
            "employees_checked": 0,
            "employees_with_probation_days": 0,
            "employees_with_contract_days": 0,
            "employees_in_range_1_20": 0,
            "sample_employees": []
        }

        # Process like check_and_send_reminders does
        for row_index, employee in enumerate(data[1:], start=1):
            if len(employee) < 12:
                continue

            debug_info["employees_checked"] += 1

            employee_name = employee[0]
            probation_remaining_days = employee[10] if len(employee) > 10 else None
            contract_remaining_days = employee[11] if len(employee) > 11 else None

            prob_days_int = None
            cont_days_int = None

            # Check probation
            if probation_remaining_days is not None and probation_remaining_days != '':
                debug_info["employees_with_probation_days"] += 1
                try:
                    prob_days_int = int(float(str(probation_remaining_days)))
                    if 1 <= prob_days_int <= 20:
                        debug_info["employees_in_range_1_20"] += 1
                        if len(debug_info["sample_employees"]) < 5:
                            debug_info["sample_employees"].append({
                                "name": employee_name,
                                "type": "Probation",
                                "days": prob_days_int
                            })
                except:
                    pass

            # Check contract
            if contract_remaining_days is not None and contract_remaining_days != '':
                debug_info["employees_with_contract_days"] += 1
                try:
                    cont_days_int = int(float(str(contract_remaining_days)))
                    if 1 <= cont_days_int <= 20:
                        debug_info["employees_in_range_1_20"] += 1
                        if len(debug_info["sample_employees"]) < 5:
                            debug_info["sample_employees"].append({
                                "name": employee_name,
                                "type": "Contract",
                                "days": cont_days_int
                            })
                except:
                    pass
        
        return jsonify({
            "success": True,
            "debug_info": debug_info
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/debug-base-fields')
def debug_base_fields():
    """Debug actual field names and data coming from Base"""
    try:
        # Get raw data from Base
        access_token = lark_client.get_access_token()
        from app import list_bitable_records
        
        response_data = list_bitable_records(
            app_token=lark_client.app_token,
            table_id=lark_client.table_id,
            access_token=access_token,
            view_id=lark_client.view_id if lark_client.view_id else None,
            page_size=5  # Just first 5 records for debugging
        )
        
        records = response_data.get('items', [])
        
        debug_info = {
            "total_records_available": len(records),
            "sample_records": []
        }
        
        for i, record in enumerate(records[:3]):  # Show first 3 records
            fields = record.get("fields", {})
            
            # Extract key fields we're looking for
            sample_record = {
                "record_number": i + 1,
                "all_field_names": list(fields.keys()),
                "key_fields": {
                    "Employee Name": fields.get('Employee Name'),
                    "Direct Leader Email": fields.get('Direct Leader Email'),
                    "Direct Leader CRM": fields.get('Direct Leader CRM'),
                    "1st Contract Renewal Date": fields.get('1st Contract Renewal Date'),
                    "Probation Period End Date": fields.get('Probation Period End Date'),
                    "Department": fields.get('Department'),
                    "Position": fields.get('Position'),
                    "Employee Status": fields.get('Employee Status'),
                    "CRM": fields.get('CRM')
                }
            }
            debug_info["sample_records"].append(sample_record)
        
        return jsonify({
            "success": True,
            "debug_info": debug_info
        })
        
    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        })


@app.route('/api/debug-old')
def debug_data_old():
    try:
        data = lark_client.get_data()
        processed_data = []
        today = datetime.now().date()
        
        for employee in data[1:]:
            if len(employee) >= 9:
                employee_info = {
                    "name": employee[0],    # Column I - Employee Name
                    "leader": employee[1] or "Manager",  # Leader Name (empty in new structure)
                    "contract_renewal": employee[2],  # Column O - Contract Renewal Date
                    "probation_end": employee[3],     # Column P - Probation Period End Date
                    "status": employee[4],            # Column AB - Employee Status
                    "leader_email_raw": employee[5],  # Column G - Leader Email
                    "leader_crm": employee[6],        # Column F - Leader CRM
                    "department": employee[7],        # Column H - Department
                    "employee_crm": employee[8],      # Column C - Employee CRM
                    "leader_email_extracted": extract_email(employee[5])
                }
                
                # Parse dates
                if employee[3]:  # Probation end date is at index 3
                    try:
                        if isinstance(employee[3], (int, float)):
                            date_obj = excel_date_to_python(employee[3])
                            employee_info["probation_end_parsed"] = date_obj.strftime('%Y-%m-%d') if date_obj else None
                            employee_info["probation_days_until"] = (date_obj.date() - today).days if date_obj else None
                        else:
                            employee_info["probation_end_parsed"] = str(employee[3])
                    except:
                        employee_info["probation_end_parsed"] = "Parse Error"
                
                processed_data.append(employee_info)
        
        return jsonify({"success": True, "data": processed_data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/all-employees-debug')
def all_employees_debug():
    """Debug endpoint to see ALL employee data including invalid emails"""
    try:
        data = lark_client.get_data()
        today = datetime.now().date()
        
        debug_data = []
        total_count = 0
        invalid_email_count = 0
        
        for employee in data[1:]:  # Skip header
            if len(employee) < 10:
                continue
                
            employee_name = employee[0]
            if not employee_name or not str(employee_name).strip():
                continue
                
            total_count += 1
            leader_email_raw = employee[6]
            leader_email = extract_email(leader_email_raw)
            
            email_valid = is_valid_email_for_sending(leader_email)
            if not email_valid:
                invalid_email_count += 1
            
            # Check evaluation dates
            probation_end = employee[3]
            contract_renewal = employee[2]
            
            probation_days = None
            contract_days = None
            
            if probation_end:
                try:
                    if isinstance(probation_end, (int, float)):
                        eval_date = excel_date_to_python(probation_end).date()
                    else:
                        eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                    probation_days = (eval_date - today).days
                except:
                    pass
            
            if contract_renewal:
                try:
                    if isinstance(contract_renewal, (int, float)):
                        eval_date = excel_date_to_python(contract_renewal).date()
                    else:
                        eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                    contract_days = (eval_date - today).days
                except:
                    pass
            
            debug_data.append({
                "employee_name": employee_name,
                "leader_email_raw": leader_email_raw,
                "leader_email_extracted": leader_email,
                "email_valid_for_sending": email_valid,
                "leader_name": employee[1] or "",
                "position": employee[5] or "",
                "department": employee[8] or "",
                "employee_status": employee[4] or "",
                "probation_end_date": employee[3] or "",
                "contract_renewal_date": employee[2] or "",
                "probation_days_remaining": probation_days,
                "contract_days_remaining": contract_days,
                "needs_probation_eval": probation_days is not None and 1 <= probation_days <= 20,
                "needs_contract_eval": contract_days is not None and 1 <= contract_days <= 20
            })
        
        return jsonify({
            "success": True,
            "total_employees": total_count,
            "employees_with_invalid_emails": invalid_email_count,
            "employees_with_valid_emails": total_count - invalid_email_count,
            "data": debug_data
        })
        
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})

# Helper function to convert timestamps to readable dates
def convert_timestamp_to_date(timestamp):
    """Convert timestamp to readable date format"""
    if not timestamp or timestamp == '':
        return 'Not specified'
    
    try:
        # Convert string timestamp to int
        timestamp_int = int(float(str(timestamp)))
        # Convert milliseconds to seconds if needed
        if timestamp_int > 9999999999:  # More than 10 digits means milliseconds
            timestamp_int = timestamp_int // 1000
        
        date_obj = datetime.fromtimestamp(timestamp_int)
        return date_obj.strftime('%Y-%m-%d')
    except:
        return str(timestamp)

# Vendor notification functionality
def get_vendor_email(contract_company):
    """Get vendor email based on contract company name"""
    if not contract_company:
        return None, None
    
    company_lower = str(contract_company).lower()
    
    if 'ضمة للاستشارات' in company_lower or 'dummah' in company_lower:
        return 'dummah@gmail.com', 'شركة ضمة للاستشارات ذات مسؤولية محدودة'
    elif 'migrate business services' in company_lower:
        return 'migrate@gmail.com', 'Migrate Business Services Co.'
    elif 'helloworld online education jordan llc' in company_lower:
        return None, 'Helloworld Online Education Jordan LLC'  # No action needed
    else:
        return None, f'Unknown Vendor (Company: {contract_company})'

def send_vendor_notification_grouped(employees_data, vendor_email, vendor_name):
    """Send vendor notification email for multiple separated employees grouped by company"""
    if not vendor_email:
        return False, "No email configured for this vendor"
    
    try:
        # Use same email configuration as reminders
        smtp_accounts = [
            {
                'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
                'smtp_port': int(os.getenv('SMTP_PORT', 587)),
                'email': os.getenv('SMTP_EMAIL'),
                'password': os.getenv('SMTP_PASSWORD')
            }
        ]
        
        # Select random SMTP account
        smtp_config = random.choice([acc for acc in smtp_accounts if acc['email'] and acc['password']])
        
        # Create email message
        message = MIMEMultipart("alternative")
        message["From"] = smtp_config['email']
        message["To"] = vendor_email
        message["Subject"] = f"Employee Separation Notification - {len(employees_data)} Employee(s)"
        
        # Create professional HTML content with table
        # Filter for October employees only
        october_employees = []
        for emp in employees_data:
            exit_date = emp.get('exit_date', '')
            if exit_date and '2024-10' in str(exit_date):
                october_employees.append(emp)

        # Create table rows with employee names and exit dates
        table_rows = ""
        for emp in october_employees:
            table_rows += f"""
                <tr>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp['name']}</td>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp.get('exit_date', 'Not specified')}</td>
                </tr>"""

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f8f9fa; }}
                .container {{ max-width: 800px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
                .header {{ text-align: center; margin-bottom: 30px; padding-bottom: 20px; border-bottom: 2px solid #007bff; }}
                .title {{ color: #007bff; font-size: 24px; font-weight: bold; margin-bottom: 10px; }}
                .company-name {{ color: #333; font-size: 18px; margin-bottom: 20px; }}
                table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                th {{ background-color: #007bff; color: white; padding: 12px; text-align: left; font-weight: bold; }}
                td {{ padding: 12px; border: 1px solid #ddd; }}
                tr:nth-child(even) {{ background-color: #f8f9fa; }}
                .footer {{ margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd; color: #666; font-size: 14px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <div class="title">Employee Separation Notification</div>
                    <div class="company-name">{vendor_name}</div>
                </div>

                <p>Dear {vendor_name},</p>

                <p>We would like to inform you that the following {len(october_employees)} employee(s) separated in October:</p>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; border: 1px solid #ddd;">
                    <thead>
                        <tr style="background-color: #007bff;">
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Employee Name</th>
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Exit Date</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows}
                    </tbody>
                </table>

                <p>Please update your records accordingly and ensure any pending transactions or access permissions are reviewed and updated.</p>

                <div class="footer">
                    <p><strong>Best regards,</strong></p>
                    <p>HR Department<br>
                    51Talk Online Education<br>
                    Date: {datetime.now().strftime('%Y-%m-%d')}</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        # Attach HTML content
        html_part = MIMEText(html_content, "html")
        message.attach(html_part)
        
        # Send email
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port']) as server:
            server.starttls(context=context)
            server.login(smtp_config['email'], smtp_config['password'])
            text = message.as_string()
            server.sendmail(smtp_config['email'], vendor_email, text)
        
        employee_names = ", ".join([emp['name'] for emp in employees_data])
        print(f"✅ Vendor notification sent to {vendor_name} ({vendor_email}) for {len(employees_data)} employees: {employee_names}")
        return True, "Email sent successfully"
        
    except Exception as e:
        print(f"❌ Failed to send vendor notification: {str(e)}")
        return False, str(e)

@app.route('/api/today-reminders')
def get_today_reminders():
    """Get today's urgent reminders for sidebar"""
    try:
        data = lark_client.get_data()
        today = datetime.now().date()
        reminders = []
        
        for employee in data[1:]:  # Skip header
            if len(employee) < 4:  # Need at least positions 0-3
                continue
                
            employee_name = employee[0]  # Position 0: Employee Name
            if not employee_name or not str(employee_name).strip():
                continue
            
            # Check probation period end date (position 3)
            probation_end_date = employee[3] if len(employee) > 3 else None
            if probation_end_date is not None and probation_end_date != '':
                try:
                    # Convert timestamp to date and calculate remaining days
                    probation_date = convert_timestamp_to_date(probation_end_date)
                    if probation_date:
                        days_remaining = (probation_date - today).days
                        if 1 <= days_remaining <= 20:
                            reminders.append({
                                'employee_name': employee_name,
                                'evaluation_type': 'Probation',
                                'days_remaining': days_remaining
                            })
                except:
                    pass

            # Check contract renewal date (position 2)
            contract_renewal_date = employee[2] if len(employee) > 2 else None
            if contract_renewal_date is not None and contract_renewal_date != '':
                try:
                    # Convert timestamp to date and calculate remaining days
                    renewal_date = convert_timestamp_to_date(contract_renewal_date)
                    if renewal_date:
                        days_remaining = (renewal_date - today).days
                        if 1 <= days_remaining <= 20:
                            reminders.append({
                                'employee_name': employee_name,
                                'evaluation_type': 'Contract',
                                'days_remaining': days_remaining
                            })
                except:
                    pass
        
        # Sort by urgency (fewer days first)
        reminders.sort(key=lambda x: x['days_remaining'])
        
        return jsonify({
            'success': True,
            'reminders': reminders[:10]  # Limit to 10 most urgent
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/check-separated-employees')
def check_separated_employees():
    """Check for employees with 'Separated' status only with optional date filtering"""
    try:
        data = lark_client.get_data()
        separated_employees = []
        
        # Get filter parameters
        date_filter = request.args.get('filter', 'all')
        custom_date = request.args.get('date', '')
        
        # Calculate date ranges for filtering
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        last_7_days = today - timedelta(days=7)
        last_30_days = today - timedelta(days=30)
        
        for employee in data[1:]:  # Skip header
            if len(employee) < 15:  # Reduced minimum requirement
                continue
                
            # Map to NEW column positions after reordering
            employee_name = employee[0] if len(employee) > 0 else None  # Employee Name - Position 0
            employee_status = employee[4] if len(employee) > 4 else None  # Employee Status - Position 4  
            exit_date = employee[28] if len(employee) > 28 else None  # Exit Date - Position 28
            exit_reason = employee[30] if len(employee) > 30 else None  # Exit Reason - Position 30
            contract_company = employee[12] if len(employee) > 12 else None  # Contract Company - Position 12
            crm = employee[9] if len(employee) > 9 else None  # Employee CRM - Position 9
            department = employee[8] if len(employee) > 8 else None  # Department - Position 8
            position = employee[5] if len(employee) > 5 else None  # Position - Position 5
            
            if not employee_name or not str(employee_name).strip():
                continue
            
            # Check for "Separated" or "Terminated" status (including misspelling "Seperated")
            if employee_status and str(employee_status).lower().strip() in ['separated', 'seperated', 'terminated']:
                exit_date_formatted = convert_timestamp_to_date(exit_date)
                
                # Apply date filtering
                if date_filter != 'all' and exit_date_formatted != '-':
                    try:
                        exit_date_obj = datetime.strptime(exit_date_formatted, '%Y-%m-%d').date()
                        
                        if date_filter == 'today' and exit_date_obj != today:
                            continue
                        elif date_filter == 'yesterday' and exit_date_obj != yesterday:
                            continue
                        elif date_filter == 'last7days' and exit_date_obj < last_7_days:
                            continue
                        elif date_filter == 'last30days' and exit_date_obj < last_30_days:
                            continue
                        elif date_filter == 'october2024' and not exit_date_formatted.startswith('2024-10'):
                            continue
                        elif date_filter == 'custom' and custom_date:
                            if exit_date_formatted != custom_date:
                                continue
                    except ValueError:
                        # Skip if date parsing fails
                        continue
                
                # Get vendor email for this contract company
                vendor_email, vendor_name = get_vendor_email(contract_company)
                
                separated_employees.append({
                    'name': employee_name,
                    'department': department or 'Not specified',
                    'crm': crm or 'Not specified',
                    'exit_reason': exit_reason or 'Not specified',
                    'exit_date': exit_date_formatted,
                    'vendor_email': vendor_email,
                    'vendor_name': vendor_name,
                    'position': position or 'Not specified',
                    'contract_company': contract_company or 'Not specified'
                })
        
        return jsonify({
            'success': True,
            'separated_employees': separated_employees,
            'filter_applied': date_filter
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/send-vendor-notifications', methods=['POST'])
def send_vendor_notifications():
    """Send vendor notifications for separated employees - grouped by company"""
    try:
        # Get separated employees
        response = check_separated_employees()
        response_data = response.get_json()
        
        if not response_data.get('success'):
            return jsonify({'success': False, 'error': 'Failed to get separated employees'}), 500
        
        separated_employees = response_data.get('separated_employees', [])
        
        # Group employees by vendor email and vendor name
        vendor_groups = {}
        for employee in separated_employees:
            vendor_email = employee.get('vendor_email')
            vendor_name = employee.get('vendor_name')
            
            if vendor_email:  # Only process if vendor email is configured
                key = f"{vendor_email}|{vendor_name}"
                if key not in vendor_groups:
                    vendor_groups[key] = {
                        'vendor_email': vendor_email,
                        'vendor_name': vendor_name,
                        'employees': []
                    }
                vendor_groups[key]['employees'].append(employee)
        
        notifications_sent = 0
        errors = []
        
        # Send one email per vendor group
        for group_key, group_data in vendor_groups.items():
            vendor_email = group_data['vendor_email']
            vendor_name = group_data['vendor_name']
            employees = group_data['employees']
            
            success, message = send_vendor_notification_grouped(employees, vendor_email, vendor_name)
            if success:
                notifications_sent += 1
            else:
                errors.append(f"Failed to notify {vendor_name}: {message}")
        
        if errors:
            return jsonify({
                'success': True,
                'notifications_sent': notifications_sent,
                'errors': errors,
                'message': f'Sent {notifications_sent} notifications with {len(errors)} errors'
            })
        else:
            return jsonify({
                'success': True,
                'notifications_sent': notifications_sent,
                'message': f'Successfully sent {notifications_sent} vendor notifications'
            })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/debug-contract-companies')
def debug_contract_companies():
    """Debug endpoint to see contract company values and employee status"""
    try:
        data = lark_client.get_data()
        contract_companies = []
        
        for i, employee in enumerate(data[1:], start=1):  # Skip header
            if len(employee) < 10:  # Reduce minimum requirement
                continue
                
            employee_name = employee[0] if len(employee) > 0 else None  # Employee Name - Position 0
            contract_company = employee[12] if len(employee) > 12 else None  # Contract Company - Position 12
            employee_status = employee[4] if len(employee) > 4 else None  # Employee Status - Position 4
            exit_reason = employee[30] if len(employee) > 30 else None  # Exit Reason - Position 30
            
            if employee_name and str(employee_name).strip():
                vendor_email, vendor_name = get_vendor_email(contract_company)
                
                contract_companies.append({
                    'row': i,
                    'employee_name': employee_name,
                    'contract_company': contract_company,
                    'employee_status': employee_status,
                    'exit_reason': exit_reason,
                    'vendor_name': vendor_name,
                    'vendor_email': vendor_email,
                    'is_separated': str(employee_status).lower().strip() in ['separated', 'seperated', 'terminated'] if employee_status else False
                })
        
        return jsonify({
            'success': True,
            'total_employees': len(contract_companies),
            'header_row': data[0] if len(data) > 0 else [],
            'contract_companies': contract_companies[:20]  # Show first 20
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/debug-data-loading')
def debug_data_loading():
    """Debug endpoint to check data loading from Base"""
    try:
        data = lark_client.get_data()
        
        return jsonify({
            'success': True,
            'total_rows': len(data),
            'header_row': data[0] if len(data) > 0 else [],
            'sample_first_5_rows': data[1:6] if len(data) > 1 else [],
            'sample_last_5_rows': data[-5:] if len(data) > 5 else [],
            'row_lengths': [len(row) for row in data[:10]] if len(data) > 0 else []
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/debug-reminder-data')
def debug_reminder_data():
    """Debug endpoint specifically for reminder data structure"""
    try:
        data = lark_client.get_data()
        
        debug_info = {
            'total_rows': len(data),
            'header_row': data[0] if len(data) > 0 else [],
            'sample_employees_with_days': []
        }
        
        # Check first 10 employees for days data
        for i, employee in enumerate(data[1:11], start=1):
            if len(employee) < 12:
                continue
                
            employee_info = {
                'row_index': i,
                'employee_name': employee[0] if len(employee) > 0 else None,
                'probation_days_col_10': employee[10] if len(employee) > 10 else None,
                'contract_days_col_11': employee[11] if len(employee) > 11 else None,
                'probation_days_col_16': employee[16] if len(employee) > 16 else None,
                'contract_days_col_18': employee[18] if len(employee) > 18 else None,
                'row_length': len(employee)
            }
            debug_info['sample_employees_with_days'].append(employee_info)
        
        return jsonify({
            'success': True,
            'debug_info': debug_info
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=False)