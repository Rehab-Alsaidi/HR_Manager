from flask import Flask, render_template, jsonify, request
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import json
import psycopg2
import os
import random
from dotenv import load_dotenv
import requests
from typing import List, Dict, Any, Optional
from database import (
    init_database, is_email_sent_today_db, mark_email_sent_db, 
    cleanup_old_email_logs_db, sync_employee_data_to_db, 
    get_employees_from_db, get_sent_emails_summary
)

# Load environment variables
load_dotenv()

# Feishu API Base Configuration
VERBOSE = False
BASE = "https://open.feishu.cn/open-apis"

class FeishuError(Exception):
    pass

def set_verbose(value: bool) -> None:
    global VERBOSE
    VERBOSE = value

# Enable verbose mode for debugging
set_verbose(False)

def debug(message: str) -> None:
    if VERBOSE:
        print(message)

def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    url = f"{BASE}/auth/v3/tenant_access_token/internal"
    debug(f"[DEBUG] Getting tenant access token for app_id: {app_id}")
    r = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=20)
    r.raise_for_status()
    data = r.json()
    debug(f"[DEBUG] Token response: {data}")
    if data.get("code") != 0:
        raise FeishuError(f"get_tenant_access_token failed: {data.get('code')} {data.get('msg')}")
    debug(f"[DEBUG] Successfully got tenant access token")
    return data["tenant_access_token"]

def list_bitable_records(app_token: str, table_id: str, access_token: str, view_id: Optional[str] = None, page_token: Optional[str] = None, page_size: int = 500) -> Dict[str, Any]:
    params: Dict[str, Any] = {'page_size': page_size}
    if view_id:
        params['view_id'] = view_id
    if page_token:
        params['page_token'] = page_token

    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"
    debug(f"[DEBUG] Listing bitable records: url={url} params={params}")
    try:
        r = requests.get(url, headers={"Authorization": f"Bearer {access_token}"}, params=params, timeout=30)
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable API error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable API error: {error_msg}")
        return resp.get('data', {})
    except requests.HTTPError as e:
        print(f"[ERROR] HTTP error while listing bitable records: {e.response.text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {e.response.text}")
    except Exception as e:
        print(f"[ERROR] Unexpected error while listing bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_delete_bitable_records(app_token: str, table_id: str, record_ids: List[str], access_token: str) -> Dict[str, Any]:
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_delete"
    payload = {'record_ids': record_ids}
    debug(f"[DEBUG] Deleting bitable records: count={len(record_ids)}")
    try:
        r = requests.post(
            url,
            json=payload,
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            timeout=30,
        )
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable delete error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable delete error: {error_msg}")
        return resp
    except requests.HTTPError as e:
        print(f"[ERROR] HTTP error while deleting bitable records: {e.response.text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {e.response.text}")
    except Exception as e:
        print(f"[ERROR] Unexpected error while deleting bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_create_bitable_records(app_token: str, table_id: str, records: List[Dict[str, Any]], access_token: str) -> Dict[str, Any]:
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_create"
    payload = {'records': records}
    debug(f"[DEBUG] Creating bitable records: count={len(records)}")
    try:
        r = requests.post(
            url,
            json=payload,
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            timeout=30,
        )
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable create error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable create error: {error_msg}")
        return resp
    except requests.HTTPError as e:
        print(f"[ERROR] HTTP error while creating bitable records: {e.response.text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {e.response.text}")
    except Exception as e:
        print(f"[ERROR] Unexpected error while creating bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

def batch_update_bitable_records(app_token: str, table_id: str, records: List[Dict[str, Any]], access_token: str) -> Dict[str, Any]:
    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records/batch_update"
    payload = {'records': records}
    debug(f"[DEBUG] Updating bitable records: count={len(records)}")
    try:
        r = requests.post(
            url,
            json=payload,
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            timeout=30,
        )
        r.raise_for_status()
        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(f"[ERROR] Bitable update error {resp.get('code')}: {error_msg}")
            raise FeishuError(f"Bitable update error: {error_msg}")
        return resp
    except requests.HTTPError as e:
        print(f"[ERROR] HTTP error while updating bitable records: {e.response.text}")
        raise FeishuError(f"HTTP error {e.response.status_code}: {e.response.text}")
    except Exception as e:
        print(f"[ERROR] Unexpected error while updating bitable records: {str(e)}")
        raise FeishuError(f"Unexpected error: {str(e)}")

# Persistent duplicate prevention system
SENT_EMAILS_LOG = "sent_emails_log.json"

def load_sent_emails_log():
    """Load the log of previously sent emails"""
    try:
        if os.path.exists(SENT_EMAILS_LOG):
            with open(SENT_EMAILS_LOG, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading sent emails log: {e}")
    return {}

def save_sent_emails_log(log_data):
    """Save the log of sent emails"""
    try:
        with open(SENT_EMAILS_LOG, 'w') as f:
            json.dump(log_data, f, indent=2)
    except Exception as e:
        print(f"Error saving sent emails log: {e}")

def is_email_already_sent_today(employee_name, leader_email, evaluation_type):
    """Check if an email for this employee was already sent today (database or file)"""
    try:
        # Try database first
        database_url = os.getenv('DATABASE_URL')
        if database_url and database_url != 'postgresql://username:password@hostname:port/database':
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
    """Mark an email as sent today (database or file)"""
    try:
        # Try database first
        database_url = os.getenv('DATABASE_URL')
        if database_url and database_url != 'postgresql://username:password@hostname:port/database':
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
    """Remove logs older than 30 days to prevent file from growing too large"""
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
        
        # Otherwise use tenant token (cached)
        if self.access_token and self.token_expires and datetime.now() < self.token_expires:
            return self.access_token
        
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
        
        # Add header row
        header_row = [
            "Employee Name", "Leader Name", "Contract Renewal Date", 
            "Probation Period End Date", "Employee Status", "Position", "Leader Email",
            "Leader CRM", "Department", "Employee CRM"
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
                            # Handle rich text fields
                            if isinstance(field_value[0], dict) and 'text' in field_value[0]:
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
                        elif field_value is not None:
                            return str(field_value)
                        return ""
                    
                    # Helper function to format dates from Base
                    def format_base_date(field_value):
                        date_str = extract_base_field_value(field_value)
                        if date_str and date_str.strip():
                            try:
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
                    
                    # Map new Base column structure
                    field_mappings = [
                        ('Employee Name', ['Employee Name']),
                        ('Leader Name', ['Direct Leader CRM']),  # Using Leader CRM as leader name for now
                        ('Contract Renewal Date', ['1st Contract Renewal Date']),
                        ('Probation Period End Date', ['Probation Period End Date']),
                        ('Employee Status', ['Employee Status']),
                        ('Position', ['Position']),
                        ('Leader Email', ['Direct Leader Email']),
                        ('Leader CRM', ['Direct Leader CRM']),
                        ('Department', ['Department']),
                        ('Employee CRM', ['CRM'])
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

    def get_sheet_data(self):
        """Get data from Lark Sheets using REST API (fallback method)"""
        token = self.get_access_token()
        
        # Use the actual internal sheet ID from the metadata response
        actual_sheet_id = "43c01e"
        
        # Read from column A to column AG to get ALL data
        import urllib.parse
        range_param = f"{actual_sheet_id}!A1:AG10000"  
        encoded_range = urllib.parse.quote(range_param, safe='')
        spreadsheet_token = os.getenv('SPREADSHEET_TOKEN')
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{encoded_range}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers)
        
        try:
            data = response.json()
            if data.get("code") == 0:
                values = data["data"]["valueRange"]["values"]
                
                # Extract columns based on sheet structure:
                extracted_data = []
                for row in values:
                    # Helper function to extract text from complex cell objects
                    def extract_text_from_cell(cell):
                        if isinstance(cell, list) and len(cell) > 0:
                            if isinstance(cell[0], dict) and 'text' in cell[0]:
                                return cell[0]['text']
                            return str(cell[0]) if cell[0] else ""
                        elif isinstance(cell, dict) and 'text' in cell:
                            return cell['text']
                        elif isinstance(cell, str):
                            return cell
                        elif cell is not None:
                            return str(cell)
                        return ""
                    
                    # Helper function to format dates
                    def format_date_if_number(value):
                        text_value = extract_text_from_cell(value)
                        if text_value and text_value.strip():
                            # Check if it's an Excel date number
                            try:
                                date_num = float(text_value)
                                if 40000 <= date_num <= 50000:  # Reasonable Excel date range
                                    # Convert Excel date to Python date
                                    epoch = datetime(1899, 12, 30)
                                    date_obj = epoch + timedelta(days=date_num)
                                    return date_obj.strftime('%Y-%m-%d')
                            except (ValueError, TypeError):
                                pass
                        return text_value
                    
                    # Always extract the row, filling missing columns with empty string
                    extracted_row = [
                        extract_text_from_cell(row[8]) if len(row) > 8 else "",   # I - Employee Name
                        "",  # Leader Name - not available in new structure, will be empty
                        format_date_if_number(row[14]) if len(row) > 14 else "",   # O - Contract Renewal Date
                        format_date_if_number(row[15]) if len(row) > 15 else "",   # P - Probation Period End Date
                        extract_text_from_cell(row[27]) if len(row) > 27 else "",   # AB - Employee Status
                        extract_text_from_cell(row[11]) if len(row) > 11 else "",   # L - Position
                        extract_text_from_cell(row[6]) if len(row) > 6 else "",    # G - Leader Email
                        extract_text_from_cell(row[5]) if len(row) > 5 else "",    # F - Leader CRM
                        extract_text_from_cell(row[7]) if len(row) > 7 else "",    # H - Department
                        extract_text_from_cell(row[2]) if len(row) > 2 else ""     # C - CRM (Employee CRM)
                    ]
                    extracted_data.append(extracted_row)
                    
                return extracted_data
            else:
                raise Exception(f"Sheet API error: {data}")
        except Exception as e:
            raise Exception(f"Failed to parse sheet response: {response.text}")
    
    def get_data(self):
        """Get data from Lark Base with fallback to Sheets"""
        try:
            data = self.get_base_data()
        except Exception as base_error:
            print(f"❌ Base access failed: {base_error}")
            print("📄 Falling back to Sheets API...")
            data = self.get_sheet_data()
        
        # Sync data to database if available
        try:
            database_url = os.getenv('DATABASE_URL')
            if database_url and database_url != 'postgresql://username:password@hostname:port/database':
                sync_employee_data_to_db(data)
        except Exception as e:
            print(f"Failed to sync data to database: {e}")
        
        return data

def get_random_email_config():
    """Get a random email configuration from the available sender emails"""
    sender_emails = os.getenv('SENDER_EMAILS', os.getenv('SENDER_EMAIL')).split(',')
    email_usernames = os.getenv('EMAIL_USERNAMES', os.getenv('EMAIL_USERNAME')).split(',')
    email_passwords = os.getenv('EMAIL_PASSWORDS', os.getenv('EMAIL_PASSWORD')).split(',')
    
    # Choose a random index
    index = random.randint(0, len(sender_emails) - 1)
    
    return {
        'sender_email': sender_emails[index].strip(),
        'username': email_usernames[index].strip(),
        'password': email_passwords[index].strip()
    }

def send_reminder_email(employee_name, leader_name, leader_email, evaluation_link, days_remaining, evaluation_type, crm_account=""):
    try:
        print(f"Attempting to send email to {leader_email} for {employee_name}")
        
        # Get random email configuration
        email_config = get_random_email_config()
        print(f"Using sender email: {email_config['sender_email']}")
        
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        with smtplib.SMTP_SSL(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT', 465)), context=context) as server:
            server.login(email_config['username'], email_config['password'])
            
            message = MIMEMultipart()
            message["Subject"] = f"Urgent: Employee Evaluation Required - {employee_name}"
            message["From"] = email_config['sender_email']
            message["To"] = leader_email
            
            # Professional email body with HTML formatting for clickable links
            html_body = f"""
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #2c5aa0; border-bottom: 2px solid #2c5aa0; padding-bottom: 10px;">
            Employee Evaluation Required
        </h2>
        
        <p>Dear <strong>{leader_name}</strong>,</p>
        
        <p>This is an urgent reminder that <strong>{employee_name}'s</strong> evaluation period is approaching and requires your immediate attention.</p>
        
        <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #2c5aa0; margin: 20px 0;">
            <h3 style="margin-top: 0; color: #2c5aa0;">Employee Details:</h3>
            <ul style="margin: 10px 0;">
                <li><strong>Name:</strong> {employee_name}</li>
                <li><strong>CRM Account:</strong> {crm_account}</li>
                <li><strong>Evaluation Type:</strong> {evaluation_type}</li>
                <li><strong>Days Remaining:</strong> <span style="color: #dc3545; font-weight: bold;">{days_remaining} days</span></li>
            </ul>
        </div>
        
        <div style="text-align: center; margin: 30px 0;">
            <a href="{evaluation_link}" 
               style="background-color: #2c5aa0; color: white; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block; font-size: 16px;"
               target="_blank">
                📋 Complete Evaluation Form
            </a>
        </div>
        
        <p style="background-color: #fff3cd; padding: 10px; border: 1px solid #ffeaa7; border-radius: 4px;">
            <strong>⚠️ Important:</strong> Please complete this evaluation before the deadline to ensure proper HR compliance.
        </p>
        
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
            server.send_message(message)
            print(f"Email sent successfully to {leader_email}")
            return True
            
    except Exception as e:
        print(f"Failed to send email to {leader_email}: {str(e)}")
        return False

def get_department_cc_emails(employees_data):
    """Get CC emails based on department mapping"""
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

def send_grouped_reminder_email(leader_name, leader_email, employees_data, evaluation_type, additional_cc_emails=None):
    """Send one email to a leader with multiple employees listed"""
    try:
        print(f"Attempting to send grouped email to {leader_email} for {len(employees_data)} employees")
        
        # Get random email configuration
        email_config = get_random_email_config()
        print(f"Using sender email: {email_config['sender_email']}")
        
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        with smtplib.SMTP_SSL(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT', 465)), context=context) as server:
            server.login(email_config['username'], email_config['password'])
            
            message = MIMEMultipart()
            message["Subject"] = f"Urgent: Multiple Employee Evaluations Required"
            message["From"] = email_config['sender_email']
            message["To"] = leader_email
            
            # Get department-based CC emails
            department_cc_emails = get_department_cc_emails(employees_data)
            
            # Add any additional CC emails if provided
            all_cc_emails = department_cc_emails.copy()
            if additional_cc_emails:
                additional_list = [email.strip() for email in additional_cc_emails.split(',') if email.strip()]
                all_cc_emails.extend(additional_list)
            
            # Remove duplicates and set CC header
            unique_cc_emails = list(set(all_cc_emails))
            if unique_cc_emails:
                message["Cc"] = ', '.join(unique_cc_emails)
            
            # Determine evaluation link and deadline
            if evaluation_type == "Probation Period Evaluation":
                evaluation_link = os.getenv('PROBATION_FORM_URL')
                form_type = "probation evaluation"
                email_intro = "Kindly find the link for the probation evaluation to be done by your side before"
                completion_note = "noting that if they pass this evaluation they will be full-time employees"
            else:
                evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                form_type = "contract renewal evaluation" 
                email_intro = "Kindly find below the name of your team employees that need to be evaluated in order to renew their contracts for a full year, in order to proceed timely it needs to be done before"
                completion_note = ""
            
            # Find the earliest deadline
            earliest_deadline = min(emp['days_remaining'] for emp in employees_data)
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
            
            server.sendmail(email_config['sender_email'], recipients, message.as_string())
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
            
        employee_name = employee[0]  # I - Employee Name
        leader_name = employee[1]  # Leader Name (empty in new structure)
        contract_renewal = employee[2]  # O - Contract Renewal Date
        probation_end = employee[3]  # P - Probation Period End Date
        employee_status = employee[4]  # AB - Employee Status
        position = employee[5]  # Position
        leader_email_raw = employee[6]  # G - Leader Email
        leader_crm = employee[7]  # F - Leader CRM
        department = employee[8]  # H - Department
        employee_crm = employee[9]  # C - Employee CRM
        
        leader_email = extract_email(leader_email_raw)
        
        # Always capture employee data, but only process for email sending if we have valid data
        if not employee_name or not str(employee_name).strip():
            continue  # Skip if no employee name
        
        # Create unique key for this employee under this leader to prevent duplicates
        employee_leader_key = f"{employee_name.strip()}|{leader_email}"
        
        # Determine which evaluation is more urgent and closer to deadline
        probation_days = None
        contract_days = None
        chosen_evaluation = None
        chosen_date = None
        chosen_days = None
        
        # Check probation end date
        if probation_end:
            try:
                if isinstance(probation_end, (int, float)):
                    eval_date = excel_date_to_python(probation_end).date()
                else:
                    eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                
                days_until = (eval_date - today).days
                
                # STRICT: Only EXACTLY 20 days
                if days_until == 20:
                    probation_days = days_until
                    if chosen_evaluation is None or days_until < chosen_days:
                        chosen_evaluation = "Probation Period Evaluation"
                        chosen_date = eval_date
                        chosen_days = days_until
            except:
                pass
        
        # Check contract renewal date  
        if contract_renewal:
            try:
                if isinstance(contract_renewal, (int, float)):
                    eval_date = excel_date_to_python(contract_renewal).date()
                else:
                    eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                
                days_until = (eval_date - today).days
                
                # STRICT: Only EXACTLY 20 days
                if days_until == 20:
                    contract_days = days_until
                    if chosen_evaluation is None or days_until < chosen_days:
                        chosen_evaluation = "Contract Renewal Evaluation"
                        chosen_date = eval_date
                        chosen_days = days_until
            except:
                pass
        
        # Only proceed if we have a valid evaluation EXACTLY at 20 days AND employee hasn't been processed yet
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
                
            print(f"Processing employee: {employee_name} | Days remaining: {chosen_days} (20 days) | Type: {chosen_evaluation}")
            processed_employees.add(employee_leader_key)
            
            # Group by leader email + evaluation type
            group_key = f"{leader_email}|{chosen_evaluation}"
            
            if group_key not in leader_groups:
                leader_groups[group_key] = {
                    'leader_name': leader_name or "Manager",
                    'leader_email': leader_email,
                    'evaluation_type': chosen_evaluation,
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
    
    # Send one email per group (leader + evaluation type)
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
            if len(employee) < 10:  # Changed back to 10 since we're only using 10 columns
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
            
            employee_leader_key = f"{employee_name.strip()}|{leader_email}"
            
            chosen_evaluation = None
            chosen_date = None
            chosen_days = None
            
            # Check probation end date
            if probation_end:
                try:
                    if isinstance(probation_end, (int, float)):
                        eval_date = excel_date_to_python(probation_end).date()
                    else:
                        eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                    
                    days_until = (eval_date - today).days
                    
                    # STRICT: Only EXACTLY 20 days
                    if days_until == 20:
                        if chosen_evaluation is None or days_until < chosen_days:
                            chosen_evaluation = "Probation Period Evaluation"
                            chosen_date = eval_date
                            chosen_days = days_until
                except:
                    pass
            
            # Check contract renewal date  
            if contract_renewal:
                try:
                    if isinstance(contract_renewal, (int, float)):
                        eval_date = excel_date_to_python(contract_renewal).date()
                    else:
                        eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                    
                    days_until = (eval_date - today).days
                    
                    # STRICT: Only EXACTLY 20 days
                    if days_until == 20:
                        if chosen_evaluation is None or days_until < chosen_days:
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_date = eval_date
                            chosen_days = days_until
                except:
                    pass
            
            # Also check for employees with "Email Sent" status (recently processed)
            email_sent_recently = (employee_status == "Email Sent")
            
            # Include if: 1) Valid evaluation in 19-25 days, OR 2) Email was sent recently (today)
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
                
                group_key = f"{leader_email}|{chosen_evaluation}"
                
                if group_key not in leader_groups:
                    leader_groups[group_key] = {
                        'leader_name': leader_name or "Manager",
                        'leader_email': leader_email,
                        'evaluation_type': chosen_evaluation,
                        'employees': [],
                        'employee_names': set(),
                        'departments': set()
                    }
                
                if employee_name not in leader_groups[group_key]['employee_names']:
                    leader_groups[group_key]['employee_names'].add(employee_name)
                    leader_groups[group_key]['departments'].add(department.strip().upper() if department else 'Unknown')
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
                'departments': ', '.join(sorted(group_data['departments'])),
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
                    
                    if days_until in [19, 20, 21, 22, 23, 24, 25]:
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
        
        # Group employees by leader email and evaluation type (EXACT same logic as check_and_send_reminders)
        leader_groups = {}
        # Track processed employees to prevent duplicates
        processed_employees = set()
        
        for row_index, employee in enumerate(data[1:], start=1):
            if len(employee) < 10:  # Changed back to 10 since we're only using 10 columns
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
            
# Debug output removed - system working correctly
            
            leader_email = extract_email(leader_email_raw)
            
            if not employee_name or not str(employee_name).strip():
                continue  # Skip if no employee name
            
            # Create unique key for this employee under this leader to prevent duplicates
            employee_leader_key = f"{employee_name.strip()}|{leader_email}"
            
            # Determine which evaluation is more urgent and closer to deadline
            chosen_evaluation = None
            chosen_date = None
            chosen_days = None
            
            # Check probation end date
            if probation_end:
                try:
                    if isinstance(probation_end, (int, float)):
                        eval_date = excel_date_to_python(probation_end).date()
                    else:
                        eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                    
                    days_until = (eval_date - today).days
                    
                    # STRICT: Only EXACTLY 20 days
                    if days_until == 20:
                        if chosen_evaluation is None or days_until < chosen_days:
                            chosen_evaluation = "Probation Period Evaluation"
                            chosen_date = eval_date
                            chosen_days = days_until
                except:
                    pass
            
            # Check contract renewal date  
            if contract_renewal:
                try:
                    if isinstance(contract_renewal, (int, float)):
                        eval_date = excel_date_to_python(contract_renewal).date()
                    else:
                        eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                    
                    days_until = (eval_date - today).days
                    
                    # STRICT: Only EXACTLY 20 days
                    if days_until == 20:
                        if chosen_evaluation is None or days_until < chosen_days:
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_date = eval_date
                            chosen_days = days_until
                except:
                    pass
            
            # Only proceed if we have a valid evaluation EXACTLY at 20 days AND employee hasn't been processed yet
            if chosen_evaluation and employee_leader_key not in processed_employees:
                # Check if we have a valid email for sending
                if not is_valid_email_for_sending(leader_email):
                    print(f"PREVIEW SKIP: {employee_name} has invalid leader email '{leader_email}'")
                    continue
                
                # CRITICAL: Check if email was already sent today to prevent duplicates in preview
                if is_email_already_sent_today(employee_name, leader_email, chosen_evaluation):
                    print(f"PREVIEW SKIP: {employee_name} - Email already sent today for {chosen_evaluation}")
                    continue
                    
                print(f"Preview: {employee_name} | Days remaining: {chosen_days} (20 days) | Type: {chosen_evaluation}")
                processed_employees.add(employee_leader_key)
                
                # Group by leader email + evaluation type
                group_key = f"{leader_email}|{chosen_evaluation}"
                
                if group_key not in leader_groups:
                    leader_groups[group_key] = {
                        'leader_name': leader_name or "Manager",
                        'leader_email': leader_email,
                        'evaluation_type': chosen_evaluation,
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
        
        # Sync to database
        from database import sync_employee_data_to_db
        sync_employee_data_to_db(data)
        
        # Get count of synced records
        from database import get_db_connection
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM employees")
        employee_count = cursor.fetchone()[0]
        cursor.close()
        conn.close()
        
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
                "needs_probation_eval": probation_days is not None and probation_days == 20,
                "needs_contract_eval": contract_days is not None and contract_days == 20
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

if __name__ == '__main__':
    app.run(debug=True, port=5002)