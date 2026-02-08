"""
HR Evaluation System - Main Flask Application.

This module contains the main Flask application for the HR Evaluation
System, handling employee evaluation reminders, vendor notifications,
and Lark/Feishu integration.
"""

import json
import os
import random
import smtplib
import ssl
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from functools import lru_cache
from typing import Any, Dict, List, Optional, Set, Tuple
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, render_template, request

from database import (
    cleanup_old_email_logs_db,
    init_database,
    is_email_sent_today_db,
    mark_email_sent_db,
)

# Load environment variables
load_dotenv()

# Feishu API Base Configuration
BASE = "https://open.feishu.cn/open-apis"

# Timezone configuration - Set to your local timezone (default: Jordan/UTC+3)
# You can override this by setting TZ environment variable (e.g., TZ="Asia/Amman")
LOCAL_TIMEZONE = ZoneInfo(os.getenv("TZ", "Asia/Amman"))

def get_local_now():
    """Get current datetime in the configured local timezone."""
    return datetime.now(LOCAL_TIMEZONE)

def get_local_today():
    """Get current date in the configured local timezone."""
    return get_local_now().date()


class FeishuError(Exception):
    """Custom exception for Feishu/Lark API errors."""

    pass


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    """
    Get tenant access token from Feishu API.

    Args:
        app_id: Lark application ID
        app_secret: Lark application secret

    Returns:
        Tenant access token string

    Raises:
        FeishuError: If token request fails or API returns error code
        requests.HTTPError: If HTTP request fails
    """
    url = f"{BASE}/auth/v3/tenant_access_token/internal"

    payload = {"app_id": app_id, "app_secret": app_secret}
    r = requests.post(url, json=payload, timeout=20)
    r.raise_for_status()

    data = r.json()

    if data.get("code") != 0:
        error_msg = (
            f"get_tenant_access_token failed: "
            f"{data.get('code')} {data.get('msg')}"
        )
        raise FeishuError(error_msg)

    return data["tenant_access_token"]

def list_bitable_records(
    app_token: str,
    table_id: str,
    access_token: str,
    view_id: Optional[str] = None,
    page_token: Optional[str] = None,
    page_size: int = 500
) -> Dict[str, Any]:
    """
    List records from a Feishu bitable with pagination support.

    Args:
        app_token: Lark Base application token
        table_id: Table ID within the Base
        access_token: Access token for authentication
        view_id: Optional view ID to filter records
        page_token: Optional pagination token for next page
        page_size: Number of records per page (default 500)

    Returns:
        Dictionary containing bitable data with items and pagination info

    Raises:
        FeishuError: If API request fails or returns error code
        requests.HTTPError: If HTTP request fails
    """
    params: Dict[str, Any] = {'page_size': page_size}
    if view_id:
        params['view_id'] = view_id
    if page_token:
        params['page_token'] = page_token

    url = f"{BASE}/bitable/v1/apps/{app_token}/tables/{table_id}/records"

    try:
        headers = {"Authorization": f"Bearer {access_token}"}
        r = requests.get(url, headers=headers, params=params, timeout=30)
        r.raise_for_status()

        resp = r.json()
        if resp.get('code') != 0:
            error_msg = resp.get('msg', 'Unknown API error')
            print(
                f"[ERROR] Bitable API error "
                f"{resp.get('code')}: {error_msg}"
            )
            raise FeishuError(f"Bitable API error: {error_msg}")

        return resp.get('data', {})

    except requests.HTTPError as e:
        error_text = e.response.text
        print(
            f"[ERROR] HTTP error while listing bitable "
            f"records: {error_text}"
        )
        raise FeishuError(
            f"HTTP error {e.response.status_code}: {error_text}"
        )

    except Exception as e:
        print(
            f"[ERROR] Unexpected error while listing bitable "
            f"records: {str(e)}"
        )
        raise FeishuError(f"Unexpected error: {str(e)}")

# Persistent duplicate prevention system
SENT_EMAILS_LOG = "sent_emails_log.json"

# Cache database availability check
_DATABASE_AVAILABLE: Optional[bool] = None


def is_database_available() -> bool:
    """
    Check if database is available (cached check).

    Returns:
        True if PostgreSQL database is available, False otherwise

    Note:
        Result is cached globally to avoid repeated connection attempts
    """
    global _DATABASE_AVAILABLE
    if _DATABASE_AVAILABLE is None:
        database_url = os.getenv('DATABASE_URL')
        _DATABASE_AVAILABLE = (
            database_url and
            'username:password@hostname' not in database_url
        )
    return _DATABASE_AVAILABLE


def load_sent_emails_log() -> Dict[str, Any]:
    """
    Load the log of previously sent emails from file.

    Returns:
        Dictionary containing sent email logs by date

    Note:
        Returns empty dict if file doesn't exist or can't be loaded
    """
    try:
        if os.path.exists(SENT_EMAILS_LOG):
            with open(SENT_EMAILS_LOG, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading sent emails log: {e}")
    return {}


def save_sent_emails_log(log_data: Dict[str, Any]) -> None:
    """
    Save the log of sent emails to file.

    Args:
        log_data: Dictionary containing sent email logs by date
    """
    try:
        with open(SENT_EMAILS_LOG, 'w') as f:
            json.dump(log_data, f, indent=2)
    except Exception as e:
        print(f"Error saving sent emails log: {e}")


def is_email_already_sent_today(
    employee_name: str,
    leader_email: str,
    evaluation_type: str
) -> bool:
    """
    Check if email for this employee was already sent today.

    Uses database if available, falls back to file-based storage.

    Args:
        employee_name: Name of the employee
        leader_email: Email address of the leader
        evaluation_type: Type of evaluation (Probation/Contract)

    Returns:
        True if email was already sent today, False otherwise
    """
    # Use cached database availability check
    if is_database_available():
        try:
            return is_email_sent_today_db(
                employee_name,
                leader_email,
                evaluation_type
            )
        except Exception as e:
            print(f"Database check failed, using file: {e}")

    # Fallback to file-based storage
    log_data = load_sent_emails_log()
    today = get_local_today().isoformat()

    # Create unique key for this employee-leader-evaluation combination
    key = f"{employee_name}|{leader_email}|{evaluation_type}"

    return log_data.get(today, {}).get(key, False)


def mark_email_as_sent(
    employee_name: str,
    leader_email: str,
    evaluation_type: str
) -> None:
    """
    Mark an email as sent today.

    Uses database if available, falls back to file-based storage.

    Args:
        employee_name: Name of the employee
        leader_email: Email address of the leader
        evaluation_type: Type of evaluation (Probation/Contract)
    """
    # Use cached database availability check
    if is_database_available():
        try:
            mark_email_sent_db(
                employee_name,
                leader_email,
                evaluation_type
            )
            return
        except Exception as e:
            print(f"Database marking failed, using file: {e}")

    # Fallback to file-based storage
    log_data = load_sent_emails_log()
    today = get_local_today().isoformat()

    if today not in log_data:
        log_data[today] = {}

    # Create unique key for this employee-leader-evaluation combination
    key = f"{employee_name}|{leader_email}|{evaluation_type}"
    log_data[today][key] = {
        "sent_at": get_local_now().isoformat(),
        "employee_name": employee_name,
        "leader_email": leader_email,
        "evaluation_type": evaluation_type
    }

    save_sent_emails_log(log_data)


def cleanup_old_logs() -> None:
    """
    Remove logs older than 30 days to prevent file from growing too large.

    Note:
        Only affects file-based storage, not database records
    """
    log_data = load_sent_emails_log()
    cutoff_date = (
        get_local_today() - timedelta(days=30)
    ).isoformat()

    # Remove entries older than 30 days
    keys_to_remove = [
        date for date in log_data.keys() if date < cutoff_date
    ]
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
    """
    Client for interacting with Lark/Feishu Base API.

    Handles authentication, data fetching, and caching for employee data
    from Lark Base. Implements 2-minute data cache for performance.

    Attributes:
        app_id: Lark application ID
        app_secret: Lark application secret
        app_token: Base application token
        table_id: Table ID within the Base
        view_id: Optional view ID for filtering
        cached_data: Cached employee data
        cache_time: Timestamp of last cache update
        cache_duration: Cache duration in seconds (default 120)
    """

    def __init__(self) -> None:
        """
        Initialize LarkClient with environment configuration.

        Raises:
            Exception: If LARK_APP_ID or LARK_APP_SECRET not configured
        """
        self.app_id = os.getenv('LARK_APP_ID')
        self.app_secret = os.getenv('LARK_APP_SECRET')

        if not self.app_id or not self.app_secret:
            raise Exception("LARK_APP_ID and LARK_APP_SECRET are required")

        # Store configuration for Base
        self.app_token = os.getenv('LARK_BASE_APP_TOKEN')
        self.table_id = os.getenv('LARK_BASE_TABLE_ID')
        self.view_id = os.getenv('LARK_BASE_VIEW_ID', '')
        self.use_user_token = (
            os.getenv('LARK_USE_USER_TOKEN', 'false').lower() == 'true'
        )

        # Initialize access token cache
        self.access_token: Optional[str] = None
        self.token_expires: Optional[datetime] = None

        # Data cache - cache for 2 minutes
        self.cached_data: Optional[List[List[Any]]] = None
        self.cache_time: Optional[float] = None
        self.cache_duration: int = 120  # seconds
    
    def get_access_token(self) -> str:
        """
        Get tenant access token for Base API.

        Returns:
            Access token string for API authentication

        Raises:
            Exception: If token cannot be obtained or user token not set
        """
        # If user token is configured, use it
        if self.use_user_token:
            user_token = os.getenv('LARK_USER_ACCESS_TOKEN')
            if user_token and user_token.strip() and user_token != '.':
                return user_token
            else:
                raise Exception(
                    "LARK_USER_ACCESS_TOKEN is required when "
                    "LARK_USE_USER_TOKEN is true"
                )

        # Always refresh tenant token to avoid invalid token errors
        # Use the new get_tenant_access_token function
        try:
            self.access_token = get_tenant_access_token(
                self.app_id,
                self.app_secret
            )
            # 2 hours minus 5 minutes buffer
            self.token_expires = (
                get_local_now() + timedelta(seconds=7200 - 300)
            )
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
            "Employee Status", "Position", "Leader Email", "Leader CRM", "Department", "2+Leader Email",
            "Employee CRM", "Probation Remaining Days", "Contract Remaining Days", "Contract Company",
            "PSID", "Big Team", "Small Team", "Marital Status", "Religion", "Joining Date",
            "2nd Contract Renewal", "Gender", "Nationality", "Birthday", "Age", "University",
            "Educational Level", "School Ranking", "Major", "Exit Date", "Exit Type", "Exit Reason",
            "Work Email address", "contract type", "service year", "Work Site", "ID N. Front", "Seperation Papers"
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
                                # Handle timestamp (milliseconds) - use local timezone
                                date_obj = datetime.fromtimestamp(field_value / 1000, tz=LOCAL_TIMEZONE)
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
                                    # Handle timestamp (milliseconds) - use local timezone
                                    date_obj = datetime.fromtimestamp(int(date_str) / 1000, tz=LOCAL_TIMEZONE)
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
                        ('2+Leader Email', ['2+Leader Email']),  # Position 9 (NEW)
                        ('Employee CRM', ['CRM']),  # Position 10
                        ('Probation Remaining Days', ['Probation Period Remaining Days']),  # Position 11 - KEEP for reminders
                        ('Contract Remaining Days', ['Remaining Limited Contract End Days']),  # Position 12 - KEEP for reminders
                        # Add new fields for vendor notifications
                        ('Contract Company', ['Specific company name for signing the employment contract']),  # Position 13
                        ('PSID', ['PSID']),  # Position 14
                        ('Big Team', ['Big Team']),  # Position 15
                        ('Small Team', ['Small Team']),  # Position 16
                        ('Marital Status', ['Marital Status']),  # Position 17
                        ('Religion', ['Religion']),  # Position 18
                        ('Joining Date', ['Joining Date']),  # Position 19
                        ('2nd Contract Renewal', ['2nd Contract Renewal']),  # Position 20
                        ('Gender', ['Gender']),  # Position 21
                        ('Nationality', ['Nationality']),  # Position 22
                        ('Birthday', ['Birthday']),  # Position 23
                        ('Age', ['Age']),  # Position 24
                        ('University', ['University']),  # Position 25
                        ('Educational Level', ['Educational Level']),  # Position 26
                        ('School Ranking', ['School Ranking']),  # Position 27
                        ('Major', ['Major']),  # Position 28
                        ('Exit Date', ['Exit Date']),  # Position 29
                        ('Exit Type', ['Exit Type']),  # Position 30
                        ('Exit Reason', ['Exit Reason']),  # Position 31
                        ('Work Email address', ['Work Email address']),  # Position 32
                        ('contract type', ['contract type']),  # Position 33
                        ('service year', ['service year']),  # Position 34
                        ('Work Site', ['Work Site']),  # Position 35
                        ('ID N. Front', ['ID N. Front']),  # Position 36
                        ('Seperation Papers', ['Seperation Papers'])  # Position 37
                    ]
                    
                    extracted_row = []
                    for field_name, field_variations in field_mappings:
                        value = ""
                        # Try each field name variation
                        for field_var in field_variations:
                            if field_var in fields:
                                if field_name in ['Contract Renewal Date', 'Probation Period End Date']:
                                    value = format_base_date(fields[field_var])
                                elif field_name == 'Seperation Papers':
                                    # Handle attachment field - store the entire structure as JSON
                                    field_value = fields[field_var]
                                    print(f"🔍 Seperation Papers field structure for record: {type(field_value)} = {field_value}")
                                    if isinstance(field_value, list) and len(field_value) > 0:
                                        # Store the entire first file object as JSON string
                                        first_file = field_value[0]
                                        if isinstance(first_file, dict):
                                            # Store the entire dict as JSON so we can extract whatever we need later
                                            import json
                                            value = json.dumps(first_file)
                                            print(f"   Stored attachment object: {value}")
                                            print(f"   Available keys: {list(first_file.keys())}")
                                        elif isinstance(first_file, str):
                                            value = first_file
                                            print(f"   Extracted from string: {value}")
                                    elif isinstance(field_value, dict):
                                        # Store the entire dict as JSON
                                        import json
                                        value = json.dumps(field_value)
                                        print(f"   Stored attachment object: {value}")
                                        print(f"   Available keys: {list(field_value.keys())}")
                                    elif isinstance(field_value, str):
                                        value = field_value
                                        print(f"   Extracted from string: {value}")
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

    
    def get_data(self, force_refresh=False):
        """Get data from Lark Base with caching"""
        # Check if cache is still valid
        if not force_refresh and self.cached_data is not None and self.cache_time is not None:
            time_since_cache = time.time() - self.cache_time
            if time_since_cache < self.cache_duration:
                print(f"📦 Using cached data ({int(self.cache_duration - time_since_cache)}s remaining)")
                return self.cached_data

        # Fetch fresh data
        data = self.get_base_data()

        # Update cache
        self.cached_data = data
        self.cache_time = time.time()

        return data

def get_random_email_config() -> Dict[str, str]:
    """
    Get email configuration from environment variables.

    Uses first email account for SMTP authentication while showing
    all HR emails in the From field for visibility.

    Returns:
        Dictionary containing:
            - sender_email: Comma-separated list of all sender emails
            - auth_email: Email to use for SMTP authentication
            - username: SMTP username
            - password: SMTP password

    Note:
        Configured via SENDER_EMAILS, EMAIL_USERNAMES, EMAIL_PASSWORDS
        environment variables
    """
    sender_emails_env = os.getenv(
        'SENDER_EMAILS',
        os.getenv('SENDER_EMAIL')
    )
    sender_emails = [
        email.strip() for email in sender_emails_env.split(',')
    ]

    usernames_env = os.getenv(
        'EMAIL_USERNAMES',
        os.getenv('EMAIL_USERNAME')
    )
    email_usernames = [
        username.strip() for username in usernames_env.split(',')
    ]

    passwords_env = os.getenv(
        'EMAIL_PASSWORDS',
        os.getenv('EMAIL_PASSWORD')
    )
    email_passwords = [
        password.strip() for password in passwords_env.split(',')
    ]

    # Use first email account to authenticate (most reliable)
    # But show all HR emails in the "From" field
    index = 0

    return {
        'sender_email': ', '.join(sender_emails),
        'auth_email': sender_emails[index],
        'username': email_usernames[index],
        'password': email_passwords[index]
    }


def get_department_cc_emails(
    employees_data: List[Dict[str, Any]]
) -> List[str]:
    """
    Get CC emails based on department mapping.

    Args:
        employees_data: List of employee dictionaries containing
                       department information

    Returns:
        List of unique CC email addresses

    Note:
        Always includes lijie14@51talk.com and department-specific
        emails based on mapping:
        - CC/GCC: wuchuan@51talk.com
        - ACC: shichuan001@51talk.com
        - EA: guanshuhao001@51talk.com, nikiyang@51talk.com
        - CM: wangjingjing@51talk.com, nikiyang@51talk.com
    """
    # Department-based CC mapping
    department_cc_mapping = {
        'CC': ['wuchuan@51talk.com'],
        'GCC': ['wuchuan@51talk.com'],
        'ACC': ['shichuan001@51talk.com'],
        'EA': ['guanshuhao001@51talk.com', 'nikiyang@51talk.com'],
        'CM': ['wangjingjing@51talk.com', 'nikiyang@51talk.com']
    }

    # Always include constant CC
    constant_cc = ['lijie14@51talk.com']

    # Collect all CC emails based on departments in this group
    cc_emails: Set[str] = set(constant_cc)

    for emp in employees_data:
        department = emp.get('department', '').strip().upper()
        if department in department_cc_mapping:
            cc_emails.update(department_cc_mapping[department])

    return list(cc_emails)

def send_grouped_reminder_email(
    leader_name: str,
    leader_email: str,
    employees_data: List[Dict[str, Any]],
    evaluation_type: str,
    additional_cc_emails: Optional[str] = None,
    second_leader_email: Optional[str] = None
) -> bool:
    """
    Send grouped reminder email to leader for multiple employees.

    Sends a single email containing information about all employees
    under a leader who need evaluations. Emails are grouped by leader,
    evaluation type, and department.

    Args:
        leader_name: Name of the team leader
        leader_email: Email address of the team leader
        employees_data: List of employee dictionaries containing:
            - name: Employee name
            - department: Department name
            - position: Job position
            - employee_crm: Employee CRM ID
            - deadline_date: Evaluation deadline (YYYY-MM-DD)
            - contract_end_date: Contract/probation end date
            - days_remaining: Days until deadline
            - leader_crm: Leader CRM ID
        evaluation_type: Type of evaluation (Probation Period
                        Evaluation/Contract Renewal Evaluation)
        additional_cc_emails: Optional comma-separated CC emails
        second_leader_email: Optional secondary leader email (2+Leader)

    Returns:
        True if email sent successfully, False otherwise

    Note:
        - Uses SMTP configuration from environment variables
        - Includes department-based CC emails automatically
        - Always includes lijie14@51talk.com in CC
        - Uses SSL for secure connection
    """
    try:
        employee_count = len(employees_data)

        # Build recipient list - include both leader emails if second leader email exists
        recipients = [leader_email]
        if second_leader_email and is_valid_email_for_sending(second_leader_email):
            recipients.append(second_leader_email)

        recipients_str = ', '.join(recipients)
        print(f"Attempting to send grouped email to {recipients_str} for {employee_count} employees")

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
            message["To"] = recipients_str
            
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
                    "to be done by your side for the following employees, "
                    "noting that if they pass this evaluation they will "
                    "be full-time employees:"
                )
            else:
                evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                email_intro = (
                    "Kindly find below the name of your team employees that "
                    "need to be evaluated in order to renew their contracts :"
                    
                )
            
            # Create employee table rows
            employee_rows = ""
            for emp in employees_data:
                employee_rows += f"""
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('position', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;"><strong>{emp['name']}</strong></td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('employee_crm', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('department', 'N/A')}</td>
                    <td style="padding: 8px; border: 1px solid #ddd; color: #dc3545; font-weight: bold;">{emp['deadline_date']}</td>
                    <td style="padding: 8px; border: 1px solid #ddd;">{emp.get('contract_end_date', 'N/A')}</td>
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

        <p>{email_intro}</p>
        
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
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Evaluation Deadline</th>
                    <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">Contract/Probation End Date</th>
                </tr>
            </thead>
            <tbody>
                {employee_rows}
            </tbody>
        </table>
        
        <div style="background-color: #fff3cd; padding: 15px; border: 1px solid #ffeaa7; border-radius: 4px; margin: 20px 0;">
            <p><strong>⚠️ Important:</strong> This evaluation must be completed within 2 working days. Please confirm by replying to this email once it's done.</p>
            <p style="margin-top: 10px;"><strong>Note:</strong> If you have already filled the evaluation for this employee, there is no need to fill it again.</p>
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

            # Get all recipients including TO and CC
            all_recipients = recipients.copy()  # Start with leader emails (TO)
            if unique_cc_emails:
                all_recipients.extend(unique_cc_emails)

            # Use auth_email for SMTP envelope, but From header shows all 3 emails
            server.sendmail(email_config['auth_email'], all_recipients, message.as_string())
            print(f"Grouped email sent successfully to {recipients_str} with CC: {', '.join(unique_cc_emails) if unique_cc_emails else 'None'}")
            return True
            
    except Exception as e:
        print(f"Failed to send grouped email to {leader_email}: {str(e)}")
        return False

def extract_email(email_data: Any) -> str:
    """
    Extract email from various data formats.

    Handles multiple data formats from Lark Base API including lists,
    dictionaries, and strings. Captures all data including '0' values.

    Args:
        email_data: Email data in various formats (list, dict, str, or None)

    Returns:
        Extracted email string, or empty string if no valid email found

    Example:
        >>> extract_email([{'text': 'user@example.com'}])
        'user@example.com'
        >>> extract_email('user@example.com')
        'user@example.com'
        >>> extract_email(['user@example.com'])
        'user@example.com'
    """
    if isinstance(email_data, list) and len(email_data) > 0:
        email_item = email_data[0]
        if isinstance(email_item, dict) and 'text' in email_item:
            return (
                str(email_item['text']).strip()
                if email_item['text'] is not None else ""
            )
        else:
            return (
                str(email_item).strip()
                if email_item is not None else ""
            )
    elif isinstance(email_data, str):
        return email_data.strip()
    elif email_data is not None:
        return str(email_data).strip()
    return ""


def is_valid_email_for_sending(email: str) -> bool:
    """
    Check if email is valid for sending.

    Validates email format and filters out placeholder values
    like '0', 'null', 'n/a', etc.

    Args:
        email: Email address to validate

    Returns:
        True if email is valid for sending, False otherwise

    Note:
        Invalid values include: '0', 'null', 'none', 'n/a', '-', 'na'
        and any string without '@' and '.'
    """
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


def excel_date_to_python(excel_date: Any) -> Optional[datetime]:
    """
    Convert Excel date number to Python datetime object.

    Args:
        excel_date: Excel date as integer or float

    Returns:
        Python datetime object, or None if not a valid Excel date

    Note:
        Excel epoch is 1900-01-01, but Excel incorrectly treats
        1900 as a leap year, so we use 1899-12-30 as epoch
    """
    if isinstance(excel_date, (int, float)):
        # Excel epoch is 1900-01-01, but Excel incorrectly
        # treats 1900 as a leap year
        epoch = datetime(1899, 12, 30)
        return epoch + timedelta(days=excel_date)
    return None


def format_date_for_display(date_value: Any) -> str:
    """
    Format date value for display.

    Args:
        date_value: Date value (Excel date number, string, or None)

    Returns:
        Formatted date string (YYYY-MM-DD) or '-' if no valid date
    """
    if isinstance(date_value, (int, float)):
        date_obj = excel_date_to_python(date_value)
        return (
            date_obj.strftime('%Y-%m-%d')
            if date_obj else str(date_value)
        )
    return str(date_value) if date_value else '-'

def check_and_send_reminders(employees_data, additional_cc_emails=None):
    sent_reminders = []
    today = get_local_today()
    
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
        leader_email_raw = employee[6]  # Leader Email (Direct Leader Email)
        leader_crm = employee[7]  # Leader CRM
        department = employee[8]  # Department
        second_leader_email_raw = employee[9] if len(employee) > 9 else None  # 2+Leader Email
        employee_crm = employee[10] if len(employee) > 10 else None  # Employee CRM
        probation_remaining_days = employee[11] if len(employee) > 11 else None  # Probation Remaining Days
        contract_remaining_days = employee[12] if len(employee) > 12 else None  # Contract Remaining Days

        leader_email = extract_email(leader_email_raw)
        second_leader_email = extract_email(second_leader_email_raw) if second_leader_email_raw else None
        
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
        
        # Variables to track both evaluation deadline and contract end date
        chosen_end_date = None

        # Check probation remaining days (from Base, minus 7 days for evaluation deadline)
        if probation_remaining_days is not None:
            try:
                days_remaining = int(float(str(probation_remaining_days)))
                # Calculate evaluation deadline (7 days before end date)
                evaluation_days_remaining = days_remaining - 7
                # Check if evaluation is due in 15-22 days (range instead of exact 20)
                if 15 <= evaluation_days_remaining <= 22:
                    chosen_evaluation = "Probation Period Evaluation"
                    chosen_days = evaluation_days_remaining
                    # Calculate the evaluation deadline date
                    chosen_date = today + timedelta(days=evaluation_days_remaining)
                    # Calculate the actual probation end date
                    chosen_end_date = today + timedelta(days=days_remaining)
            except:
                pass

        # Check contract remaining days (from Base, minus 7 days for evaluation deadline)
        if contract_remaining_days is not None:
            try:
                days_remaining = int(float(str(contract_remaining_days)))
                # Calculate evaluation deadline (7 days before end date)
                evaluation_days_remaining = days_remaining - 7
                # Check if evaluation is due in 15-22 days (range instead of exact 20)
                if 15 <= evaluation_days_remaining <= 22:
                    # If probation is also in range, probation takes priority
                    if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                        chosen_evaluation = "Contract Renewal Evaluation"
                        chosen_days = evaluation_days_remaining
                        # Calculate the evaluation deadline date
                        chosen_date = today + timedelta(days=evaluation_days_remaining)
                        # Calculate the actual contract end date
                        chosen_end_date = today + timedelta(days=days_remaining)
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
                    'second_leader_email': second_leader_email,  # Add 2+Leader Email
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
                    'contract_end_date': chosen_end_date.strftime('%Y-%m-%d') if chosen_end_date else 'N/A',
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
            additional_cc_emails,
            group_data.get('second_leader_email')  # Pass 2+Leader Email
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
def initialize_app() -> None:
    """
    Initialize the application with database setup.

    Attempts to initialize PostgreSQL database if DATABASE_URL is
    configured. Falls back to file-based storage if database is
    not available.

    Note:
        Called automatically on application startup
        Performs cleanup of old email logs (30+ days)
    """
    try:
        # Try to initialize database if DATABASE_URL is available
        database_url = os.getenv('DATABASE_URL')
        if database_url:
            print("🔧 Initializing database...")
            init_database()
            cleanup_old_email_logs_db()
            print("✅ Database initialized successfully")
        else:
            print(
                "📝 Using file-based storage "
                "(DATABASE_URL not configured)"
            )
            cleanup_old_logs()
    except Exception as e:
        print(
            f"⚠️  Database initialization failed, "
            f"using file-based storage: {e}"
        )
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
        today = get_local_today()
        
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
            second_leader_email_raw = employee[9] if len(employee) > 9 else None  # 2+Leader Email
            employee_crm = employee[10] if len(employee) > 10 else None

            leader_email = extract_email(leader_email_raw)
            second_leader_email = extract_email(second_leader_email_raw) if second_leader_email_raw else None

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
            probation_remaining_days = employee[11] if len(employee) > 11 else None
            contract_remaining_days = employee[12] if len(employee) > 12 else None

            # Check probation remaining days (from Base, minus 7 days for evaluation deadline)
            if probation_remaining_days is not None:
                try:
                    days_remaining = int(float(str(probation_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        chosen_evaluation = "Probation Period Evaluation"
                        chosen_days = evaluation_days_remaining
                        # Calculate the evaluation deadline date
                        chosen_date = today + timedelta(days=evaluation_days_remaining)
                except:
                    pass

            # Check contract remaining days (from Base, minus 7 days for evaluation deadline)
            if contract_remaining_days is not None:
                try:
                    days_remaining = int(float(str(contract_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        # If probation is also in range, probation takes priority
                        if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_days = evaluation_days_remaining
                            # Calculate the evaluation deadline date
                            chosen_date = today + timedelta(days=evaluation_days_remaining)
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
                        'second_leader_email': second_leader_email,
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
                'second_leader_email': group_data.get('second_leader_email'),
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
    """Welcome dashboard with quick access cards"""
    try:
        data = lark_client.get_data()
        total_employees = 0

        # Count total employees
        if data:
            for employee in data[1:]:
                if len(employee) >= 9:
                    employee_name = str(employee[0]).strip() if employee[0] else ""
                    if employee_name and employee_name not in ['-', 'null', 'None', '']:
                        total_employees += 1

        return render_template('index.html', total_employees=total_employees)
    except Exception as e:
        return f"Error loading dashboard: {str(e)}", 500

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

@app.route('/api/preview-reminders')
def preview_reminders():
    """Preview which emails would be sent without actually sending them"""
    try:
        data = lark_client.get_data()
# Debug output removed - system working correctly
        today = get_local_today()
        
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
            second_leader_email_raw = employee[9] if len(employee) > 9 else None  # 2+Leader Email
            employee_crm = employee[10] if len(employee) > 10 else None
            probation_remaining_days = employee[11] if len(employee) > 11 else None
            contract_remaining_days = employee[12] if len(employee) > 12 else None

            leader_email = extract_email(leader_email_raw)
            second_leader_email = extract_email(second_leader_email_raw) if second_leader_email_raw else None

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

            # Check probation remaining days (from Base, minus 7 days for evaluation deadline)
            if probation_remaining_days is not None:
                try:
                    days_remaining = int(float(str(probation_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        chosen_evaluation = "Probation Period Evaluation"
                        chosen_days = evaluation_days_remaining
                        # Calculate the evaluation deadline date
                        chosen_date = today + timedelta(days=evaluation_days_remaining)
                except:
                    pass

            # Check contract remaining days (from Base, minus 7 days for evaluation deadline)
            if contract_remaining_days is not None:
                try:
                    days_remaining = int(float(str(contract_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        # If probation is also in range, probation takes priority
                        if chosen_evaluation is None or chosen_evaluation != "Probation Period Evaluation":
                            chosen_evaluation = "Contract Renewal Evaluation"
                            chosen_days = evaluation_days_remaining
                            # Calculate the evaluation deadline date
                            chosen_date = today + timedelta(days=evaluation_days_remaining)
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

# Helper function to convert timestamps to readable dates
def convert_timestamp_to_date(timestamp: Any) -> str:
    """
    Convert timestamp to readable date format.

    Args:
        timestamp: Timestamp value (int, float, or string)

    Returns:
        Formatted date string (YYYY-MM-DD) or 'Not specified'
        if timestamp is empty or invalid

    Note:
        Handles both seconds and milliseconds timestamps
        (converts milliseconds to seconds automatically)
    """
    if not timestamp or timestamp == '':
        return 'Not specified'

    try:
        # Convert string timestamp to int
        timestamp_int = int(float(str(timestamp)))
        # Convert milliseconds to seconds if needed
        # More than 10 digits means milliseconds
        if timestamp_int > 9999999999:
            timestamp_int = timestamp_int // 1000

        # Use configured local timezone instead of UTC to avoid day offset issues
        date_obj = datetime.fromtimestamp(timestamp_int, tz=LOCAL_TIMEZONE)
        formatted_date = date_obj.strftime('%Y-%m-%d')
        print(f"🕐 Timestamp conversion: {timestamp} → {timestamp_int} → {formatted_date} ({LOCAL_TIMEZONE})")
        return formatted_date
    except Exception as e:
        print(f"❌ Failed to convert timestamp {timestamp}: {str(e)}")
        return str(timestamp)


# Vendor notification functionality
def get_vendor_email(
    contract_company: Optional[str]
) -> Tuple[Optional[str], Optional[str]]:
    """
    Get vendor email and name based on contract company.

    Args:
        contract_company: Contract company name from employee data

    Returns:
        Tuple of (vendor_email, vendor_name)
        Returns (None, None) if contract_company is None

    Note:
        Current mappings:
        - Dummah/ضمة للاستشارات -> alsaidirehab@51talk.com
        - Migrate Business Services -> alsaidirehab@51talk.com
        - Helloworld Online Education -> No action needed (None)
        - Unknown company -> None
    """
    if not contract_company:
        return None, None

    company_lower = str(contract_company).lower()

    # Both Dummah and Migrate Business Services send to alsaidirehab
    if 'ضمة للاستشارات' in company_lower or 'dummah' in company_lower:
        return (
            'alsaidirehab@51talk.com',
            'شركة ضمة للاستشارات ذات مسؤولية محدودة'
        )
    elif 'migrate business services' in company_lower:
        return (
            'alsaidirehab@51talk.com',
            'Migrate Business Services Co.'
        )
    elif 'helloworld online education jordan llc' in company_lower:
        # No action needed
        return None, 'Helloworld Online Education Jordan LLC'
    else:
        return None, f'Unknown Vendor (Company: {contract_company})'

def download_lark_attachment_by_url(download_url: str, access_token: Optional[str] = None) -> Optional[Tuple[bytes, str]]:
    """
    Download file from Lark/Feishu using download URL.

    Args:
        download_url: Download URL from attachment field (can be 'url' or 'tmp_url')
        access_token: Optional Lark access token for authentication

    Returns:
        Tuple of (file_data, filename) or None if download fails
    """
    try:
        print(f"📥 Downloading from URL: {download_url[:100]}...")

        # Add authorization header if access token is provided
        headers = {}
        if access_token:
            headers["Authorization"] = f"Bearer {access_token}"

        response = requests.get(download_url, headers=headers, timeout=30)
        response.raise_for_status()

        # Get filename from Content-Disposition header or URL
        filename = "attachment.pdf"  # default
        if 'Content-Disposition' in response.headers:
            content_disp = response.headers['Content-Disposition']
            # Parse Content-Disposition header properly
            # Example: attachment; filename="file.pdf"; filename*=UTF-8''file.pdf
            if 'filename=' in content_disp:
                # Extract the first filename value
                parts = content_disp.split('filename=')
                if len(parts) > 1:
                    # Get the value after filename=
                    filename_part = parts[1].split(';')[0].strip('"').strip("'")
                    if filename_part:
                        filename = filename_part
        else:
            # Try to extract from URL
            from urllib.parse import urlparse, unquote
            parsed = urlparse(download_url)
            path_parts = parsed.path.split('/')
            if path_parts:
                filename = unquote(path_parts[-1]) or "attachment.pdf"

        print(f"   ✅ Successfully downloaded: {filename} ({len(response.content)} bytes)")
        return (response.content, filename)
    except Exception as e:
        print(f"❌ Failed to download from URL: {str(e)}")
        try:
            print(f"   Response status: {response.status_code}")
            print(f"   Response body: {response.text[:200]}")
        except:
            pass
        return None

def download_lark_attachment(file_token: str, access_token: str) -> Optional[Tuple[bytes, str]]:
    """
    Download file from Lark/Feishu using file token.

    Args:
        file_token: Lark file token from attachment field
        access_token: Lark access token for authentication

    Returns:
        Tuple of (file_data, filename) or None if download fails
    """
    try:
        # Try the files endpoint first (for attachments)
        url = f"{BASE}/drive/v1/files/{file_token}/download"
        headers = {"Authorization": f"Bearer {access_token}"}

        print(f"📥 Attempting to download file: {file_token}")
        print(f"   Using URL: {url}")

        response = requests.get(url, headers=headers, timeout=30)

        # If files endpoint fails, try medias endpoint
        if response.status_code == 400 or response.status_code == 404:
            print(f"   Files endpoint failed, trying medias endpoint...")
            url = f"{BASE}/drive/v1/medias/{file_token}/download"
            response = requests.get(url, headers=headers, timeout=30)

        response.raise_for_status()

        # Get filename from Content-Disposition header
        filename = "attachment.pdf"  # default
        if 'Content-Disposition' in response.headers:
            content_disp = response.headers['Content-Disposition']
            if 'filename=' in content_disp:
                filename = content_disp.split('filename=')[1].strip('"')

        print(f"   ✅ Successfully downloaded: {filename}")
        return (response.content, filename)
    except Exception as e:
        print(f"❌ Failed to download Lark attachment {file_token}: {str(e)}")
        try:
            print(f"   Response status: {response.status_code}")
            print(f"   Response body: {response.text[:200]}")
        except:
            pass
        return None

def send_vendor_notification_grouped(
    employees_data: List[Dict[str, Any]],
    vendor_email: str,
    vendor_name: str
) -> Tuple[bool, str]:
    """
    Send vendor notification email for separated employees.

    Sends a single email to vendor containing all separated employees
    from that company, grouped by contract company.

    Args:
        employees_data: List of employee dictionaries containing:
            - name: Employee name
            - national_id: National ID (ID N. Front)
            - exit_date: Exit date (YYYY-MM-DD)
            - exit_type: Exit type (Forced/Voluntary)
            - exit_reason: Exit reason (additional details)
        vendor_email: Vendor email address
        vendor_name: Vendor company name

    Returns:
        Tuple of (success, message):
            - success: True if email sent successfully, False otherwise
            - message: Success or error message

    Note:
        - Includes CC to lijie14@51talk.com (HR)
        - Uses same SMTP configuration as reminder emails
        - Email contains table with: employee name, national ID,
          exit date, exit type (Forced/Voluntary)
    """
    if not vendor_email:
        return False, "No email configured for this vendor"
    
    try:
        # Get email configuration (uses same config as reminders)
        email_config = get_random_email_config()

        smtp_config = {
            'smtp_server': os.getenv('SMTP_SERVER'),
            'smtp_port': int(os.getenv('SMTP_PORT', 465)),
            'email': email_config['auth_email'],
            'password': email_config['password'],
            'sender_email': email_config['sender_email']
        }
        
        # Create email message
        message = MIMEMultipart("alternative")
        message["From"] = smtp_config['sender_email']
        message["To"] = vendor_email
        message["Cc"] = "lijie14@51talk.com"
        message["Subject"] = f"Employee Separation Notification - {len(employees_data)} Employee(s)"
        
        # Create professional HTML content with table
        # Use all employees passed to this function (already filtered by date)

        # Create table rows with employee names, exit dates, exit type, national ID, and separation papers
        table_rows = ""
        attachment_list = []  # Track attachments for summary section
        for emp in employees_data:
            # Parse separation papers to show user-friendly text
            separation_display = "-"
            separation_papers = emp.get('separation_papers', '-')

            if separation_papers and separation_papers not in ['-', '', 'Not specified']:
                try:
                    import json
                    attachment_obj = json.loads(separation_papers)
                    # Show the filename if available with professional styling
                    if 'name' in attachment_obj:
                        attachment_filename = attachment_obj['name']
                        safe_emp_name = emp['name'].replace(' ', '_')
                        full_attachment_name = f"{safe_emp_name}_{attachment_filename}"

                        # Add to attachment list for summary
                        attachment_list.append({
                            'employee': emp['name'],
                            'filename': full_attachment_name
                        })

                        # Professional display with icon - reference attachment number
                        attachment_number = len(attachment_list)
                        separation_display = f'''
                            <div style="display: flex; align-items: center; gap: 6px;">
                                <span style="background: #007bff; color: white; border-radius: 50%; width: 20px; height: 20px; display: inline-flex; align-items: center; justify-content: center; font-size: 11px; font-weight: bold;">{attachment_number}</span>
                                <div style="color: #007bff; font-weight: 600; font-size: 13px;">{attachment_filename}</div>
                            </div>
                        '''
                    else:
                        separation_display = "✓ Attached (See Attachments)"
                except (json.JSONDecodeError, TypeError):
                    # If not JSON, might be plain text filename
                    separation_display = "✓ Attached (See Attachments)"

            table_rows += f"""
                <tr>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp['name']}</td>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp.get('national_id', '-')}</td>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp.get('exit_date', 'Not specified')}</td>
                    <td style="padding: 12px; border: 1px solid #ddd;">{emp.get('exit_type', '-')}</td>
                    <td style="padding: 12px; border: 1px solid #ddd;">{separation_display}</td>
                </tr>"""

        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; margin: 0; padding: 20px; background-color: #f8f9fa; }}
                .container {{ max-width: 900px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
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

                <p>We would like to inform you that the following {len(employees_data)} employee(s) have separated from the company:</p>

                <table style="width: 100%; border-collapse: collapse; margin: 20px 0; border: 1px solid #ddd;">
                    <thead>
                        <tr style="background-color: #007bff;">
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Employee Name</th>
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">National ID</th>
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Exit Date</th>
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Exit Type</th>
                            <th style="padding: 12px; border: 1px solid #ddd; text-align: left; color: white;">Separation Papers</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows}
                    </tbody>
                </table>

                <p>Please update your records accordingly and ensure any pending transactions or access permissions are reviewed and updated.</p>

                {'<div style="margin-top: 30px; padding: 20px; background: #f8f9fa; border-left: 4px solid #007bff; border-radius: 4px;"><h3 style="color: #007bff; margin-top: 0; font-size: 16px; margin-bottom: 15px;">📎 Attached Documents (' + str(len(attachment_list)) + ')</h3>' + ''.join([f'<div style="padding: 10px; margin: 8px 0; background: white; border-radius: 4px; border: 1px solid #dee2e6; display: flex; align-items: center; gap: 10px;"><span style="background: #007bff; color: white; border-radius: 50%; width: 24px; height: 24px; display: inline-flex; align-items: center; justify-content: center; font-size: 12px; font-weight: bold; flex-shrink: 0;">{i+1}</span><div><div style="font-weight: 600; color: #333; font-size: 13px;">{att["filename"]}</div><div style="color: #6c757d; font-size: 11px;">Employee: {att["employee"]}</div></div></div>' for i, att in enumerate(attachment_list)]) + '<p style="margin-top: 15px; margin-bottom: 0; color: #6c757d; font-size: 12px; font-style: italic;">💡 Tip: Scroll down to view and download attached documents</p></div>' if attachment_list else ''}

                <div class="footer">
                    <p><strong>Best regards,</strong></p>
                    <p>HR Department<br>
                    51Talk Online Education<br>
                    Date: {get_local_now().strftime('%Y-%m-%d')}</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        # Attach HTML content
        html_part = MIMEText(html_content, "html")
        message.attach(html_part)

        # Attach separation papers for each employee if available
        attachments_count = 0
        try:
            import json
            access_token = lark_client.get_access_token()
            print(f"\n📎 Processing attachments for {len(employees_data)} employees...")

            for idx, emp in enumerate(employees_data, 1):
                separation_papers = emp.get('separation_papers')
                print(f"\n[{idx}/{len(employees_data)}] Employee: {emp['name']}")
                print(f"    Separation papers value: {separation_papers}")

                if separation_papers and separation_papers not in ['-', '', 'Not specified']:
                    file_data = None

                    # Try to parse as JSON first (if it's an attachment object)
                    try:
                        attachment_obj = json.loads(separation_papers)
                        print(f"    📎 Parsed attachment object - has {len(attachment_obj)} fields")

                        # Try 'url' field first (direct download link with permissions)
                        if 'url' in attachment_obj and attachment_obj['url']:
                            print(f"    → Using 'url' field for download")
                            file_data = download_lark_attachment_by_url(attachment_obj['url'], access_token)
                        # Try tmp_url (direct download link)
                        elif 'tmp_url' in attachment_obj and attachment_obj['tmp_url']:
                            print(f"    → Using tmp_url for download")
                            file_data = download_lark_attachment_by_url(attachment_obj['tmp_url'], access_token)
                        # Otherwise try file_token
                        elif 'file_token' in attachment_obj and attachment_obj['file_token']:
                            print(f"    → Using file_token for download: {attachment_obj['file_token']}")
                            file_data = download_lark_attachment(attachment_obj['file_token'], access_token)
                        # Try other possible token fields
                        elif 'token' in attachment_obj and attachment_obj['token']:
                            print(f"    → Using token for download: {attachment_obj['token']}")
                            file_data = download_lark_attachment(attachment_obj['token'], access_token)

                    except (json.JSONDecodeError, TypeError) as e:
                        # Not JSON, treat as simple file token string
                        print(f"    Not JSON (error: {e}), treating as file token")
                        if isinstance(separation_papers, str) and separation_papers.strip():
                            file_token = separation_papers.strip()
                            print(f"    → Processing file token: {file_token}")
                            file_data = download_lark_attachment(file_token, access_token)

                    if file_data:
                        file_content, filename = file_data

                        # Create attachment with proper MIME type
                        # Use application/pdf for PDF files, otherwise octet-stream
                        mime_type = 'application/pdf' if filename.lower().endswith('.pdf') else 'application/octet-stream'
                        part = MIMEBase('application', mime_type.split('/')[-1])
                        part.set_payload(file_content)
                        encoders.encode_base64(part)

                        # Add header with employee name prefix
                        safe_emp_name = emp['name'].replace(' ', '_')
                        attachment_name = f"{safe_emp_name}_{filename}"

                        # Add Content-Disposition header for attachment
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename="{attachment_name}"'
                        )

                        message.attach(part)
                        attachments_count += 1
                        print(f"    ✅ Attached: {attachment_name} (Total attachments: {attachments_count})")
                    else:
                        print(f"    ⚠️  Could not download separation papers")
                else:
                    print(f"    ℹ️  No separation papers to attach")

            print(f"\n✅ Total attachments added to email: {attachments_count}")

        except Exception as e:
            print(f"  ⚠️  Error attaching separation papers: {str(e)}")
            import traceback
            traceback.print_exc()

        # Send email using SSL (same as reminders)
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE

        with smtplib.SMTP_SSL(smtp_config['smtp_server'], smtp_config['smtp_port'], context=context) as server:
            server.login(email_config['username'], smtp_config['password'])
            server.send_message(message)
        
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
        today = get_local_today()
        reminders = []

        for employee in data[1:]:  # Skip header
            if len(employee) < 13:  # Need at least up to remaining days fields
                continue

            employee_name = employee[0]  # Position 0: Employee Name
            employee_status = employee[4] if len(employee) > 4 else None  # Employee Status

            if not employee_name or not str(employee_name).strip():
                continue

            # Only process Active employees
            if not employee_status or str(employee_status).strip().lower() != 'active':
                continue

            probation_remaining_days = employee[11] if len(employee) > 11 else None
            contract_remaining_days = employee[12] if len(employee) > 12 else None

            # Check probation remaining days (from Base, minus 7 days for evaluation deadline)
            if probation_remaining_days is not None:
                try:
                    days_remaining = int(float(str(probation_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        reminders.append({
                            'employee_name': employee_name,
                            'evaluation_type': 'Probation',
                            'days_remaining': evaluation_days_remaining
                        })
                except:
                    pass

            # Check contract remaining days (from Base, minus 7 days for evaluation deadline)
            if contract_remaining_days is not None:
                try:
                    days_remaining = int(float(str(contract_remaining_days)))
                    # Calculate evaluation deadline (7 days before end date)
                    evaluation_days_remaining = days_remaining - 7
                    # Check if evaluation is due in 15-22 days (range instead of exact 20)
                    if 15 <= evaluation_days_remaining <= 22:
                        reminders.append({
                            'employee_name': employee_name,
                            'evaluation_type': 'Contract',
                            'days_remaining': evaluation_days_remaining
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

        # Normalize custom date to YYYY-MM-DD format
        normalized_custom_date = ''
        if custom_date:
            try:
                # Try parsing various date formats - prioritize DD-MM-YYYY for ambiguous dates
                for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y']:
                    try:
                        parsed_date = datetime.strptime(custom_date, fmt).date()
                        normalized_custom_date = parsed_date.strftime('%Y-%m-%d')
                        print(f"📅 Parsed custom date '{custom_date}' as {normalized_custom_date} using format {fmt}")
                        break
                    except ValueError:
                        continue

                if not normalized_custom_date:
                    print(f"❌ Could not parse custom date: {custom_date}")
            except Exception as e:
                print(f"Could not parse custom date '{custom_date}': {e}")

        # Calculate date ranges for filtering
        today = get_local_today()
        yesterday = today - timedelta(days=1)
        last_7_days = today - timedelta(days=7)
        last_30_days = today - timedelta(days=30)
        
        for employee in data[1:]:  # Skip header
            if len(employee) < 15:  # Reduced minimum requirement
                continue
                
            # Map to NEW column positions after reordering
            employee_name = employee[0] if len(employee) > 0 else None  # Employee Name - Position 0
            employee_status = employee[4] if len(employee) > 4 else None  # Employee Status - Position 4
            exit_date = employee[29] if len(employee) > 29 else None  # Exit Date - Position 29
            exit_type = employee[30] if len(employee) > 30 else None  # Exit Type - Position 30 (Forced/Voluntary)
            exit_reason = employee[31] if len(employee) > 31 else None  # Exit Reason - Position 31
            contract_company = employee[13] if len(employee) > 13 else None  # Contract Company - Position 13
            crm = employee[10] if len(employee) > 10 else None  # Employee CRM - Position 10
            department = employee[8] if len(employee) > 8 else None  # Department - Position 8
            position = employee[5] if len(employee) > 5 else None  # Position - Position 5
            national_id = employee[36] if len(employee) > 36 else None  # ID N. Front - Position 36
            separation_papers = employee[37] if len(employee) > 37 else None  # Seperation Papers - Position 37

            if not employee_name or not str(employee_name).strip():
                continue

            # Check for "Separated" or "Terminated" status (including misspelling "Seperated")
            if employee_status and str(employee_status).lower().strip() in ['separated', 'seperated', 'terminated']:
                exit_date_formatted = convert_timestamp_to_date(exit_date)

                # Debug: Print employee separation info
                print(f"🔍 Found separated employee: {employee_name} | Status: {employee_status} | Exit Date: {exit_date} | Formatted: {exit_date_formatted}")

                # Skip employees without exit dates (don't show "Not specified" or "-" in vendor notifications)
                if not exit_date_formatted or exit_date_formatted in ['-', 'Not specified', '']:
                    print(f"  ⚠️  Skipping {employee_name} - No valid exit date")
                    continue

                # Apply date filtering
                if exit_date_formatted != '-':
                    try:
                        exit_date_obj = datetime.strptime(exit_date_formatted, '%Y-%m-%d').date()

                        if date_filter == 'today':
                            # Only show if exit date is exactly today
                            print(f"  📅 Checking today filter: exit_date={exit_date_obj}, today={today}, match={exit_date_obj == today}")
                            if exit_date_obj != today:
                                continue
                        elif date_filter == 'last7days':
                            # Show if within last 7 days (including today)
                            if exit_date_obj < last_7_days or exit_date_obj > today:
                                continue
                        elif date_filter == 'yesterday':
                            print(f"  📅 Checking yesterday filter: exit_date={exit_date_obj}, yesterday={yesterday}, match={exit_date_obj == yesterday}")
                            if exit_date_obj != yesterday:
                                print(f"    ❌ Not yesterday, skipping")
                                continue
                            print(f"    ✅ Matched yesterday!")
                        elif date_filter == 'last30days' and exit_date_obj < last_30_days:
                            continue
                        elif date_filter == 'october2024' and not exit_date_formatted.startswith('2024-10'):
                            continue
                        elif date_filter == 'custom' and normalized_custom_date:
                            if exit_date_formatted != normalized_custom_date:
                                continue
                        # For 'all' filter, show all dates
                    except ValueError:
                        # If date parsing fails, still include the employee (date might be invalid format)
                        pass

                # Get vendor email for this contract company
                vendor_email, vendor_name = get_vendor_email(contract_company)

                # Add ALL separated employees (not just vendor companies)
                separated_employees.append({
                    'name': employee_name,
                    'department': department or '-',
                    'crm': crm or '-',
                    'exit_type': exit_type or '-',
                    'exit_reason': exit_reason or '-',
                    'exit_date': exit_date_formatted,
                    'vendor_email': vendor_email if vendor_email else 'No action needed',
                    'vendor_name': vendor_name,
                    'position': position or '-',
                    'contract_company': contract_company or 'N/A',
                    'national_id': national_id or '-',
                    'separation_papers': separation_papers or '-'
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
        # Get filter parameters from POST request
        request_data = request.get_json() if request.is_json else {}
        date_filter = request_data.get('filter', 'all')
        custom_date = request_data.get('date', '')

        # Normalize custom date to YYYY-MM-DD format
        normalized_custom_date = ''
        if custom_date:
            try:
                # Try parsing various date formats - prioritize DD-MM-YYYY for ambiguous dates
                for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y']:
                    try:
                        parsed_date = datetime.strptime(custom_date, fmt).date()
                        normalized_custom_date = parsed_date.strftime('%Y-%m-%d')
                        print(f"📅 Parsed custom date '{custom_date}' as {normalized_custom_date} using format {fmt}")
                        break
                    except ValueError:
                        continue

                if not normalized_custom_date:
                    print(f"❌ Could not parse custom date: {custom_date}")
            except Exception as e:
                print(f"Could not parse custom date '{custom_date}': {e}")

        # Get separated employees with the same filter that was used in the UI
        data = lark_client.get_data()
        separated_employees = []

        # Calculate date ranges for filtering
        today = get_local_today()
        yesterday = today - timedelta(days=1)
        last_7_days = today - timedelta(days=7)
        last_30_days = today - timedelta(days=30)

        for employee in data[1:]:  # Skip header
            if len(employee) < 15:
                continue

            # Map to NEW column positions
            employee_name = employee[0] if len(employee) > 0 else None
            employee_status = employee[4] if len(employee) > 4 else None
            exit_date = employee[29] if len(employee) > 29 else None
            exit_type = employee[30] if len(employee) > 30 else None  # Exit Type - Position 30 (Forced/Voluntary)
            exit_reason = employee[31] if len(employee) > 31 else None
            contract_company = employee[13] if len(employee) > 13 else None
            crm = employee[10] if len(employee) > 10 else None
            department = employee[8] if len(employee) > 8 else None
            position = employee[5] if len(employee) > 5 else None
            national_id = employee[36] if len(employee) > 36 else None  # ID N. Front - Position 36
            separation_papers = employee[37] if len(employee) > 37 else None  # Seperation Papers - Position 37

            if not employee_name or not str(employee_name).strip():
                continue

            # Check for "Separated" or "Terminated" status
            if employee_status and str(employee_status).lower().strip() in ['separated', 'seperated', 'terminated']:
                exit_date_formatted = convert_timestamp_to_date(exit_date)

                # Skip employees without exit dates (don't show "Not specified" or "-" in vendor notifications)
                if not exit_date_formatted or exit_date_formatted in ['-', 'Not specified', '']:
                    continue

                # Apply date filtering - SAME LOGIC as check_separated_employees
                if exit_date_formatted != '-':
                    try:
                        exit_date_obj = datetime.strptime(exit_date_formatted, '%Y-%m-%d').date()

                        if date_filter == 'today':
                            if exit_date_obj != today:
                                continue
                        elif date_filter == 'last7days':
                            if exit_date_obj < last_7_days or exit_date_obj > today:
                                continue
                        elif date_filter == 'yesterday' and exit_date_obj != yesterday:
                            continue
                        elif date_filter == 'last30days' and exit_date_obj < last_30_days:
                            continue
                        elif date_filter == 'october2024' and not exit_date_formatted.startswith('2024-10'):
                            continue
                        elif date_filter == 'custom' and normalized_custom_date:
                            if exit_date_formatted != normalized_custom_date:
                                continue
                    except ValueError:
                        pass

                # Get vendor email for this contract company
                vendor_email, vendor_name = get_vendor_email(contract_company)

                # Add ALL separated employees (not just vendor companies)
                separated_employees.append({
                    'name': employee_name,
                    'department': department or '-',
                    'crm': crm or '-',
                    'exit_type': exit_type or '-',
                    'exit_reason': exit_reason or '-',
                    'exit_date': exit_date_formatted,
                    'vendor_email': vendor_email if vendor_email else 'No action needed',
                    'vendor_name': vendor_name,
                    'position': position or '-',
                    'contract_company': contract_company or 'N/A',
                    'national_id': national_id or '-',
                    'separation_papers': separation_papers or '-'
                })

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

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=False)