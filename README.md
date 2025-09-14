# HR Manager

A Flask-based HR evaluation management system that automatically sends reminder emails for employee evaluations (probation periods and contract renewals).

## Features

- **Automated Email Reminders**: Sends email notifications to team leaders 20 days before evaluation deadlines
- **Lark/Feishu Integration**: Fetches employee data from Lark spreadsheets
- **Dual Evaluation Types**: Handles both probation period and contract renewal evaluations
- **Web Dashboard**: View all employees and their evaluation status
- **PostgreSQL Database**: Stores employee data and evaluation history
- **Railway Deployment**: Ready for one-click deployment on Railway

## Tech Stack

- **Backend**: Flask (Python)
- **Database**: PostgreSQL
- **Email**: SMTP integration
- **External APIs**: Lark/Feishu API
- **Deployment**: Railway

## Setup

### Environment Variables

Create a `.env` file with the following variables:

```env
# Lark (Feishu) Configuration
LARK_APP_ID=your_lark_app_id
LARK_APP_SECRET=your_lark_app_secret

# Email Configuration
SMTP_SERVER=your_smtp_server
SMTP_PORT=465
SENDER_EMAIL=your_sender_email
EMAIL_USERNAME=your_email_username
EMAIL_PASSWORD=your_email_password

# Sheet Configuration
SPREADSHEET_TOKEN=your_spreadsheet_token
SHEET_ID=your_sheet_id

# Database Configuration
DATABASE_URL=postgresql://username:password@hostname:port/database

# Evaluation Form Links
PROBATION_FORM_URL=your_probation_form_url
CONTRACT_RENEWAL_FORM_URL=your_contract_renewal_form_url
```

### Local Development

1. Clone the repository:
   ```bash
   git clone https://github.com/Rehab-Alsaidi/HR_Manager.git
   cd HR_Manager
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up environment variables (create `.env` file)

4. Run the application:
   ```bash
   python app.py
   ```

5. Create database schema:
   ```python
   from app import create_schema
   create_schema()
   ```

## Railway Deployment

1. Connect your GitHub repository to Railway
2. Add a PostgreSQL service
3. Set all environment variables in Railway dashboard
4. Deploy automatically with `railway.json`

### Post-Deployment Setup

After deployment, create the database schema by calling the `create_schema()` function once.

## API Endpoints

- `GET /` - Main dashboard showing all employees
- `GET /api/employees` - Get all employee data
- `GET /api/send-reminders` - Send reminder emails for upcoming evaluations
- `GET /api/debug` - Debug endpoint for data inspection

## Database Schema

The application creates three main tables:

- **employees**: Store employee information and evaluation dates
- **evaluation_reminders**: Track sent reminders
- **evaluation_forms**: Store evaluation form URLs

## How It Works

1. **Data Fetching**: Retrieves employee data from Lark spreadsheet
2. **Date Processing**: Converts Excel date formats to Python datetime
3. **Reminder Logic**: Identifies employees with evaluations due within 20 days
4. **Email Sending**: Sends HTML/text emails to team leaders
5. **Status Tracking**: Updates employee status to prevent duplicate emails

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

This project is private and proprietary.