# HR Evaluation System - Deployment Guide

## Local Development Setup

### Prerequisites
- Python 3.9+
- Git

### Local Environment Setup

1. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure Environment Variables**
   - The `.env` file is already configured for local development
   - Make sure these are set:
     ```
     SMTP_SERVER=smtp.qiye.aliyun.com
     SMTP_PORT=465
     SENDER_EMAILS=sarakhateeb@51talk.com,ruba.hr@51talk.com
     EMAIL_USERNAMES=sarakhateeb@51talk.com,ruba.hr@51talk.com
     EMAIL_PASSWORDS=OTCySfjutTyzkLei,6RuXMLlkj4qWRP24
     ```

3. **Run Locally**
   ```bash
   python app.py
   ```
   - Access at: `http://localhost:5002`
   - Uses file-based storage (`sent_emails_log.json`)

---

## Railway Deployment

### First-Time Deployment

1. **Push Code to GitHub**
   ```bash
   git add .
   git commit -m "Update configuration"
   git push origin main
   ```

2. **Railway Setup**
   - Go to [Railway.app](https://railway.app)
   - Create new project from your GitHub repo
   - Railway will auto-detect the Python app

3. **Add Environment Variables in Railway Dashboard**

   Required variables:
   ```
   LARK_APP_ID=cli_a8345d86b838900d
   LARK_APP_SECRET=x0lqIDixh7GLAxgfDTmcub53CcMGOUzw

   SMTP_SERVER=smtp.qiye.aliyun.com
   SMTP_PORT=465

   SENDER_EMAILS=sarakhateeb@51talk.com,ruba.hr@51talk.com
   EMAIL_USERNAMES=sarakhateeb@51talk.com,ruba.hr@51talk.com
   EMAIL_PASSWORDS=OTCySfjutTyzkLei,6RuXMLlkj4qWRP24

   LARK_BASE_APP_TOKEN=TXoMbOZ2kayp3eswfQFczHDrnYb
   LARK_BASE_TABLE_ID=tbllMHIbIKBqzffq
   LARK_BASE_VIEW_ID=vew7eShPsQ

   PROBATION_FORM_URL=https://iu8uoujh41.feishu.cn/share/base/form/shrcnjqsZG30PeJZdYoPbGVLUyd
   CONTRACT_RENEWAL_FORM_URL=https://iu8uoujh41.feishu.cn/share/base/form/shrcn0UagkT0a0k06maC4GklGEe
   ```

4. **Add PostgreSQL Database (Optional but Recommended)**
   - In Railway, click "New" → "Database" → "PostgreSQL"
   - Railway automatically sets `DATABASE_URL`
   - Database will be used instead of file-based storage

5. **Deploy**
   - Railway auto-deploys when you push to GitHub
   - Get your app URL from Railway dashboard

---

## Updating the Application

### For Local Changes
```bash
# Make your changes
python app.py  # Test locally
```

### Deploy to Railway
```bash
git add .
git commit -m "Your update message"
git push origin main
```
Railway will automatically redeploy.

---

## Environment Comparison

| Feature | Local Development | Railway Production |
|---------|------------------|-------------------|
| **Storage** | File (`sent_emails_log.json`) | PostgreSQL Database |
| **SMTP** | Lark/Feishu SMTP (Alibaba) | Lark/Feishu SMTP (Alibaba) |
| **Port** | 5002 | Auto-assigned by Railway |
| **Database** | Not required | PostgreSQL (recommended) |

---

## Troubleshooting

### Local Issues

**SMTP Error: "please run connect() first"**
- Check `.env` has `SMTP_SERVER` and `SMTP_PORT`
- Verify email passwords are correct

**Database errors locally**
- Ignore - local uses file-based storage
- Make sure `DATABASE_URL` line is commented out in `.env`

### Railway Issues

**App not starting**
- Check Railway logs
- Verify all environment variables are set
- Ensure PostgreSQL database is added

**Emails not sending**
- Check SMTP credentials in Railway environment variables
- Verify Lark/Feishu email accounts have SMTP enabled
- Ensure email passwords are correct app-specific passwords

---

## Configuration Files

- `.env` - Local environment variables (DO NOT commit secrets)
- `requirements.txt` - Python dependencies
- `Procfile` - Railway deployment configuration
- `runtime.txt` - Python version specification

---

## Support

For issues:
1. Check Railway logs
2. Test locally first
3. Verify environment variables match between local and Railway
