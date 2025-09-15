from flask import Flask, render_template, jsonify
import requests
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import json
import psycopg2
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)

class LarkClient:
    def __init__(self):
        self.access_token = None
        self.token_expires = None
    
    def get_access_token(self):
        if self.access_token and self.token_expires and datetime.now() < self.token_expires:
            return self.access_token
        
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        payload = {
            "app_id": os.getenv('LARK_APP_ID'),
            "app_secret": os.getenv('LARK_APP_SECRET')
        }
        
        response = requests.post(url, json=payload)
        try:
            data = response.json()
        except Exception as e:
            raise Exception(f"Failed to parse token response as JSON: {response.text}")
        
        if data.get("code") == 0:
            self.access_token = data["tenant_access_token"]
            self.token_expires = datetime.now() + timedelta(seconds=data["expire"] - 300)
            return self.access_token
        else:
            raise Exception(f"Failed to get access token: {data}")
    
    def get_sheet_data(self):
        token = self.get_access_token()
        
        # Use the actual internal sheet ID from the metadata response
        actual_sheet_id = "43c01e"  # This is the real sheetId from the API response
        
        # Use the correct "read single range" API endpoint from documentation
        # Read from column A to column AC to get all data including our needed columns
        # E=Leader Name, I=Employee Name, Q=Contract Renewal Date, R=Probation Period End Date, Z=Employee Status, AC=Work Email
        import urllib.parse
        range_param = f"{actual_sheet_id}!A1:AC1000"  # Read full range to get all columns
        encoded_range = urllib.parse.quote(range_param, safe='')
        url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{os.getenv('SPREADSHEET_TOKEN')}/values/{encoded_range}"
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers)
        
        try:
            data = response.json()
            if data.get("code") == 0:
                values = data["data"]["valueRange"]["values"]
                
                # Extract columns based on actual sheet structure
                # A=Employee Name, B=Leader Name, C=Probation End, D=Contract Renewal, E=Status, F=Leader Email
                extracted_data = []
                for row in values:
                    # Always extract the row, filling missing columns with empty string
                    extracted_row = [
                        row[1] if len(row) > 1 else "",   # B - Leader Name
                        row[0] if len(row) > 0 else "",   # A - Employee Name
                        row[3] if len(row) > 3 else "",   # D - Contract Renewal Date
                        row[2] if len(row) > 2 else "",   # C - Probation Period End Date
                        row[4] if len(row) > 4 else "",   # E - Employee Status
                        row[5] if len(row) > 5 else ""    # F - Leader Email
                    ]
                    extracted_data.append(extracted_row)
                    
                return extracted_data
            else:
                raise Exception(f"API error: {data}")
        except Exception:
            raise Exception(f"Failed to parse response: {response.text}")

def send_reminder_email(employee_name, leader_name, leader_email, evaluation_link, days_remaining, evaluation_type):
    try:
        print(f"Attempting to send email to {leader_email} for {employee_name}")
        
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        
        with smtplib.SMTP_SSL(os.getenv('SMTP_SERVER'), int(os.getenv('SMTP_PORT', 465)), context=context) as server:
            server.login(os.getenv('EMAIL_USERNAME'), os.getenv('EMAIL_PASSWORD'))
            
            message = MIMEMultipart()
            message["Subject"] = f"Urgent: Employee Evaluation Required - {employee_name}"
            message["From"] = os.getenv('SENDER_EMAIL')
            message["To"] = leader_email
            
            # Professional email body with proper greeting
            body = f"""
Dear {leader_name},

This is an urgent reminder that {employee_name}'s evaluation period is approaching and requires your immediate attention.

Employee Details:
- Name: {employee_name}
- Evaluation Type: {evaluation_type}
- Days Remaining: {days_remaining} days

Please complete the evaluation using this link:
linkkkk

It is important to complete this evaluation before the deadline to ensure proper HR compliance.

Thank you for your prompt attention to this matter.

Best regards,
HR Team
51Talk
            """
            
            message.attach(MIMEText(body, "plain"))
            server.send_message(message)
            print(f"Email sent successfully to {leader_email}")
            return True
            
    except Exception as e:
        print(f"Failed to send email to {leader_email}: {str(e)}")
        return False

def extract_email(email_data):
    if isinstance(email_data, list) and len(email_data) > 0:
        email_item = email_data[0]
        if isinstance(email_item, dict) and 'text' in email_item:
            return email_item['text']
    elif isinstance(email_data, str):
        return email_data
    return None

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

def check_and_send_reminders(employees_data):
    sent_reminders = []
    today = datetime.now().date()
    
    for row_index, employee in enumerate(employees_data[1:], start=1):
        if len(employee) < 6:  # Now we only need 6 columns
            continue
            
        leader_name = employee[0]  # Column E - Leader Name
        employee_name = employee[1]  # Column I - Employee Name  
        contract_renewal = employee[2]  # Column Q - Contract Renewal Date
        probation_end = employee[3]  # Column R - Probation Period End Date
        status = employee[4]  # Column Z - Employee Status
        leader_email_raw = employee[5]  # Column AC - Work Email
        
        leader_email = extract_email(leader_email_raw)
        
        if not leader_email:
            continue
        
        # Check probation end date
        if probation_end:
            try:
                if isinstance(probation_end, (int, float)):
                    eval_date = excel_date_to_python(probation_end).date()
                else:
                    eval_date = datetime.strptime(str(probation_end), "%Y-%m-%d").date()
                
                days_until = (eval_date - today).days
                
                if days_until <= 20 and days_until >= 0:
                    # Use static probation evaluation link
                    evaluation_link = os.getenv('PROBATION_FORM_URL')
                    email_sent = send_reminder_email(employee_name, leader_name, leader_email, evaluation_link, days_until, "Probation Period Evaluation")
                    if email_sent:
                        # Update status in the original data
                        employees_data[row_index][4] = "Email Sent"  # Status is now at index 4
                        sent_reminders.append({
                            "employee": employee_name,
                            "leader": leader_name,
                            "type": "Probation Period",
                            "days": days_until,
                            "email_sent": True
                        })
            except:
                continue
        
        # Check contract renewal date  
        if contract_renewal:
            try:
                if isinstance(contract_renewal, (int, float)):
                    eval_date = excel_date_to_python(contract_renewal).date()
                else:
                    eval_date = datetime.strptime(str(contract_renewal), "%Y-%m-%d").date()
                
                days_until = (eval_date - today).days
                
                if days_until <= 20 and days_until >= 0:
                    # Use static contract renewal evaluation link
                    evaluation_link = os.getenv('CONTRACT_RENEWAL_FORM_URL')
                    email_sent = send_reminder_email(employee_name, leader_name, leader_email, evaluation_link, days_until, "Contract Renewal Evaluation")
                    if email_sent:
                        # Update status in the original data
                        employees_data[row_index][4] = "Email Sent"  # Status is now at index 4
                        sent_reminders.append({
                            "employee": employee_name,
                            "leader": leader_name,
                            "type": "Contract Renewal",
                            "days": days_until,
                            "email_sent": True
                        })
            except:
                continue
    
    return sent_reminders

lark_client = LarkClient()

@app.route('/')
def index():
    try:
        data = lark_client.get_sheet_data()
        # Process data to format dates for display
        formatted_data = []
        total_employees = 0
        
        if data:
            # Keep header row as is
            formatted_data.append(data[0])
            
            # Format employee data and count actual employees
            for employee in data[1:]:
                if len(employee) >= 6:
                    # Check if this is a real employee record (has employee name)
                    if employee[1] and str(employee[1]).strip():  # Employee name exists
                        total_employees += 1
                        
                        # Determine which evaluation link to show (probation vs contract renewal)
                        # For display, we'll show the probation form by default, but both will be available
                        evaluation_link = os.getenv('PROBATION_FORM_URL')  # Default to probation form
                        
                        formatted_employee = [
                            employee[1],  # Employee Name (Column I)
                            employee[0],  # Leader Name (Column E)
                            format_date_for_display(employee[3]),  # Probation End Date (Column R)
                            format_date_for_display(employee[2]),  # Contract Renewal Date (Column Q)
                            evaluation_link,  # Static Evaluation Link (probation form)
                            employee[4],  # Status (Column Z)
                            employee[5]   # Leader Email (Column AC)
                        ]
                        formatted_data.append(formatted_employee)
        
        return render_template('index.html', employees=formatted_data, total_employees=total_employees)
    except Exception as e:
        return f"Error fetching data: {str(e)}", 500

@app.route('/api/employees')
def api_employees():
    try:
        data = lark_client.get_sheet_data()
        return jsonify({"success": True, "data": data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/send-reminders')
def send_reminders():
    try:
        data = lark_client.get_sheet_data()
        sent = check_and_send_reminders(data)
        return jsonify({"success": True, "sent_reminders": sent})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route('/api/debug')
def debug_data():
    try:
        data = lark_client.get_sheet_data()
        processed_data = []
        today = datetime.now().date()
        
        for employee in data[1:]:
            if len(employee) >= 6:
                employee_info = {
                    "leader": employee[0],  # Column E - Leader Name
                    "name": employee[1],    # Column I - Employee Name
                    "contract_renewal": employee[2],  # Column Q - Contract Renewal Date
                    "probation_end": employee[3],     # Column R - Probation Period End Date
                    "status": employee[4],            # Column Z - Employee Status
                    "leader_email_raw": employee[5],  # Column AC - Work Email
                    "leader_email_extracted": extract_email(employee[5])
                }
                
                # Parse dates
                if employee[3]:  # Probation end date is now at index 3
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

if __name__ == '__main__':
    app.run(debug=True, port=5001)