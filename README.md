# HR Evaluation System

> **Enterprise-grade employee evaluation management system with automated reminders, vendor notifications, and Lark/Feishu integration.**

[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![Flask](https://img.shields.io/badge/flask-2.3.3-green.svg)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/license-Proprietary-red.svg)](LICENSE)

---

## 📋 Overview

A comprehensive HR management solution that automates employee evaluation tracking, sends timely reminders to team leaders, and manages vendor notifications for separated employees. Built with Flask and optimized for Railway deployment.

### Key Features

- ✅ **Automated Email Reminders** - Smart notification system for probation and contract renewal evaluations
- ✅ **Vendor Notifications** - Automated separation notifications to external vendors
- ✅ **Lark/Feishu Integration** - Real-time employee data synchronization
- ✅ **Intelligent Caching** - 2-minute data cache for 85-99% faster page loads
- ✅ **Duplicate Prevention** - PostgreSQL-backed email tracking with file-based fallback
- ✅ **Department-Based CC** - Automatic email routing based on department
- ✅ **Modern Dashboard** - Clean, responsive web interface
- ✅ **Production Ready** - Optimized for Railway deployment with gunicorn

---

## 🚀 Quick Start

### Prerequisites

- Python 3.11 or higher
- PostgreSQL (optional, file-based fallback available)
- Lark/Feishu account with API access
- SMTP credentials (Lark/Feishu email)

### Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/HR_Eval.git
cd HR_Eval

# Install dependencies
pip install -r requirements.txt

# Configure environment
cp .env.example .env
# Edit .env with your credentials

# Run the application
python app.py
```

Access the application at `http://localhost:5002`

---

## ⚙️ Configuration

### Environment Variables

Create a `.env` file in the project root:

```env
# Lark/Feishu API Configuration
LARK_APP_ID=your_app_id
LARK_APP_SECRET=your_app_secret
LARK_BASE_APP_TOKEN=your_base_token
LARK_BASE_TABLE_ID=your_table_id
LARK_BASE_VIEW_ID=your_view_id

# SMTP Configuration (Lark/Feishu Email)
SMTP_SERVER=smtp.qiye.aliyun.com
SMTP_PORT=465
SENDER_EMAILS=email1@company.com,email2@company.com
EMAIL_USERNAMES=email1@company.com,email2@company.com
EMAIL_PASSWORDS=password1,password2

# Evaluation Form URLs
PROBATION_FORM_URL=https://your-probation-form-url
CONTRACT_RENEWAL_FORM_URL=https://your-contract-renewal-url

# Database (Optional - uses file storage if not set)
DATABASE_URL=postgresql://user:pass@host:port/db
```

### Department CC Email Mapping

Configured in `app.py`:
- **CC/GCC**: wuchuan@51talk.com
- **ACC**: shichuan001@51talk.com
- **EA**: guanshuhao001@51talk.com, nikiyang@51talk.com
- **CM**: wangjingjing@51talk.com, nikiyang@51talk.com
- **All**: lijie14@51talk.com (constant CC)

---

## 📊 Application Structure

### Pages

#### 1. **Dashboard** (`/`)
- Welcome page with system overview
- Quick access cards to main features
- Real-time statistics

#### 2. **Today's Reminders** (`/reminders`)
- View pending evaluation reminders
- Grouped by leader, department, and evaluation type
- One-click email sending with preview
- Shows 2+Leader email recipients

#### 3. **Vendor Notifications** (`/vendor-notifications`)
- Manage separated employee notifications
- Date-based filtering (today, last 7/30 days, custom)
- Sends to vendor emails with HR CC
- Excludes employees without exit dates

### API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Main dashboard |
| `/reminders` | GET | Today's reminders page |
| `/vendor-notifications` | GET | Vendor notifications page |
| `/api/today-reminders` | GET | Get reminder data (JSON) |
| `/api/send-reminders` | POST | Send all reminder emails |
| `/api/check-separated-employees` | GET | Get separated employees |
| `/api/send-vendor-notifications` | POST | Send vendor notifications |

---

## 🏗️ Architecture

### Technology Stack

| Component | Technology | Purpose |
|-----------|------------|---------|
| **Backend** | Flask 2.3.3 | Web framework |
| **Database** | PostgreSQL / File-based | Email tracking & persistence |
| **Caching** | In-memory | 2-minute data cache |
| **Email** | SMTP (Alibaba) | Lark/Feishu email server |
| **API Integration** | Lark/Feishu API | Employee data source |
| **Web Server** | Gunicorn | Production WSGI server |
| **Deployment** | Railway | Cloud platform |

### Data Flow

```
Lark/Feishu Base
       ↓
   [Cache Layer]
       ↓
   Flask Routes
       ↓
 Business Logic
    ↙     ↘
Email      Database
System     (PostgreSQL/File)
```

### Performance Optimizations

- **2-minute data cache**: 85-99% faster page loads
- **Cached database checks**: Eliminates repeated connection attempts
- **Grouped email sending**: Single email per leader/department/type
- **Lazy loading**: Resources allocated on demand

---

## 📧 Email System

### Reminder Emails

**Triggered when:**
- Probation end date within 19-25 days
- Contract renewal date within 19-25 days

**Recipients:**
- Direct Leader Email (To)
- 2+Leader Email (To, if exists)
- Department CC emails (CC)
- lijie14@51talk.com (CC - constant)

**Features:**
- Grouped by leader, evaluation type, and department
- HTML formatted with clickable form links
- Includes employee details and deadline countdown
- Duplicate prevention (won't resend same day)
- Note about not refilling if already completed

### Vendor Notifications

**Triggered for:**
- Separated/Terminated employees
- Filtered by date range

**Recipients:**
- Vendor email (To): alsaidirehab@51talk.com
  - Migrate Business Services Co.
  - شركة ضمة للاستشارات ذات مسؤولية محدودة
- HR copy (CC): lijie14@51talk.com

**Includes:**
- Employee name
- Exit date
- Exit type/reason

---

## 🗄️ Database Schema

### Email Tracking Table

```sql
CREATE TABLE sent_emails (
    id SERIAL PRIMARY KEY,
    employee_name VARCHAR(255) NOT NULL,
    leader_email VARCHAR(255) NOT NULL,
    evaluation_type VARCHAR(100) NOT NULL,
    sent_date DATE NOT NULL,
    sent_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(employee_name, leader_email, evaluation_type, sent_date)
);
```

### Fallback Storage

If PostgreSQL is not available:
- Uses `sent_emails_log.json` (file-based)
- Same duplicate prevention logic
- Auto-cleanup of old entries (30 days)

---

## 🚢 Deployment

### Railway Deployment

1. **Connect Repository**
   ```bash
   # Push to GitHub
   git push origin main
   ```

2. **Configure Railway**
   - Connect GitHub repository
   - Add PostgreSQL database (optional)
   - Set environment variables

3. **Environment Variables** (Railway Dashboard)
   ```
   LARK_APP_ID, LARK_APP_SECRET
   SMTP_SERVER, SMTP_PORT
   SENDER_EMAILS, EMAIL_USERNAMES, EMAIL_PASSWORDS
   PROBATION_FORM_URL, CONTRACT_RENEWAL_FORM_URL
   LARK_BASE_APP_TOKEN, LARK_BASE_TABLE_ID, LARK_BASE_VIEW_ID
   ```

4. **Deploy**
   - Railway auto-deploys on push
   - Database tables created automatically

### Local Development

```bash
# Run with Flask development server
python app.py

# Run with production server (gunicorn)
gunicorn app:app --bind 0.0.0.0:5002 --timeout 120
```

---

## 📈 Performance Metrics

| Metric | Before Optimization | After Optimization |
|--------|-------------------|-------------------|
| **Dashboard Load** | 5-8 seconds | < 1 second ⚡ |
| **Reminders Page** | 6-10 seconds | < 1 second ⚡ |
| **Vendor Page** | 5-7 seconds | < 1 second ⚡ |
| **Subsequent Loads** | Same | < 50ms ⚡ |
| **Cache Hit Rate** | 0% | 95%+ |

---

## 🔒 Security

- ✅ Environment variables for sensitive data
- ✅ `.env` excluded from version control
- ✅ PostgreSQL password encryption
- ✅ SMTP SSL/TLS encryption
- ✅ Input validation and sanitization
- ✅ Duplicate prevention system

---

## 📝 Documentation

- **[DEPLOYMENT.md](DEPLOYMENT.md)** - Complete deployment guide
- **[PERFORMANCE_OPTIMIZATIONS.md](PERFORMANCE_OPTIMIZATIONS.md)** - Performance details
- **[CLEANUP_SUMMARY.md](CLEANUP_SUMMARY.md)** - Code cleanup documentation
- **[DEPLOYMENT_CONFIG_CHECK.md](DEPLOYMENT_CONFIG_CHECK.md)** - Configuration verification

---

## 🛠️ Development

### Code Standards

- **PEP 8** compliance
- **Type hints** for all functions
- **Docstrings** for documentation
- **Comments** for complex logic

### Project Structure

```
HR_Eval/
├── app.py                    # Main Flask application
├── database.py               # Database operations
├── templates/                # HTML templates
│   ├── index.html           # Dashboard
│   ├── reminders.html       # Reminders page
│   └── vendor_notifications.html
├── .env                      # Environment variables (not in git)
├── requirements.txt          # Python dependencies
├── Procfile                  # Railway configuration
├── runtime.txt              # Python version
└── README.md                # This file
```

### Dependencies

```
Flask==2.3.3              # Web framework
requests==2.31.0          # HTTP library
psycopg2-binary==2.9.7    # PostgreSQL adapter
python-dotenv==1.0.0      # Environment variables
gunicorn==21.2.0          # Production server
```

---

## 🐛 Troubleshooting

### Common Issues

**Email not sending:**
- Check SMTP credentials in `.env`
- Verify Lark/Feishu email accounts are active
- Check Railway environment variables

**Database connection errors:**
- Normal if `DATABASE_URL` not set (uses file storage)
- Verify PostgreSQL is running (Railway)
- Check database credentials

**Slow performance:**
- Check cache is working (console shows "📦 Using cached data")
- Verify Lark API credentials
- Check network connectivity

**"Not specified" in vendor emails:**
- System now filters out employees without exit dates
- Only employees with valid exit dates are included

---

## 📊 System Requirements

### Minimum

- Python 3.11+
- 256MB RAM
- 100MB disk space

### Recommended

- Python 3.11.9
- 512MB RAM
- PostgreSQL database
- Railway deployment

---

## 🤝 Contributing

This is a proprietary system for internal use. For bugs or feature requests, contact the development team.

---

## 📄 License

Proprietary - All rights reserved. For internal company use only.

---

## 👥 Support

For technical support or questions:
- Create an issue in the repository
- Contact: HR Development Team

---

## 🔄 Changelog

### v2.0.0 (2024-12-23)
- ✨ Added 2-minute data caching (85-99% faster)
- ✨ Added vendor notifications with date filtering
- ✨ Added 2+Leader email support
- ✨ Added department-based CC routing
- ✨ Redesigned dashboard with 2-card layout
- 🐛 Fixed vendor email filtering for separated employees
- 🐛 Fixed database connection optimization
- 🗑️ Removed unused code and templates
- 📝 Improved documentation

### v1.0.0 (2024-09-14)
- 🎉 Initial release
- ✨ Basic reminder system
- ✨ Lark integration
- ✨ PostgreSQL support

---

**Built with ❤️ for efficient HR management**
