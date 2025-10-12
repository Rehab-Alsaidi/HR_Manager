# How to Switch from Lark Sheets to Lark Base

## Quick Setup (5 Minutes)

### Step 1: Update Your .env File
Open `/Users/user/Desktop/HR_Eval/.env` and change this line:

**Change from:**
```
LARK_DATA_SOURCE_TYPE=sheet
```

**Change to:**
```
LARK_DATA_SOURCE_TYPE=base
```

### Step 2: Verify Base Configuration
Make sure these lines are already in your `.env` file (they should be there):
```
LARK_BASE_APP_TOKEN=Koj9b1Z31aDl9ysTU1dcHus5n2c
LARK_BASE_TABLE_ID=tblTSv6G3XBYEEF7
LARK_BASE_VIEW_ID=vewEjey2Cn
```

### Step 3: Fix App Permissions ⚠️ **MOST IMPORTANT STEP**
Your Lark app needs permission to read Base data:

1. Go to [Lark Developer Console](https://open.feishu.cn/app)
2. Find your app (`cli_a8345d86b838900d`) 
3. Click **Permissions & Scopes**
4. Add these permissions:
   - ✅ `bitable:read` (Read base data)
   - ✅ `bitable:record` (Read base records)
5. **IMPORTANT**: Click "Submit for Review" or "Apply" to activate permissions
6. Wait 5-10 minutes for permissions to propagate

### Step 4: Restart the Application
```bash
# Stop current app (Ctrl+C if running)
# Then restart:
python3 app.py
```

## Expected Base Field Names

Your Lark Base table should have these exact field names:

| Field Name | Type | Description |
|------------|------|-------------|
| `Employee Name` | Text | Full employee name |
| `Leader Name` | Text | Manager name (optional) |
| `Contract Renewal Date` | Date | Contract renewal date |
| `Probation Period End Date` | Date | Probation end date |
| `Employee Status` | Select/Text | Current status |
| `Leader Email` | Email | Manager's email |
| `Leader CRM` | Text | Manager's CRM ID |
| `Department` | Text | Employee department |
| `Employee CRM` | Text | Employee CRM ID |

## Troubleshooting

### Issue: "Forbidden" Error
**Solution:** Your app doesn't have Base permissions
- Follow Step 3 above to add `bitable:read` permission
- Wait 5-10 minutes for permissions to take effect

### Issue: No Data Returned
**Solution:** Check field names match exactly (case-sensitive)
- Field names must match the table above exactly
- Use the Base URL you provided to verify field names

### Issue: App Falls Back to Sheets
**Solution:** This is normal if Base access fails
- The app will automatically use Sheet data if Base fails
- Check the console for any error messages

## Reverting to Sheets

If Base doesn't work, you can always go back to Sheets:

1. Change `.env` back to:
   ```
   LARK_DATA_SOURCE_TYPE=sheet
   ```
2. Restart the app

## Base URL Reference
Your Base URL: `https://iu8uoujh41.feishu.cn/base/Koj9b1Z31aDl9ysTU1dcHus5n2c?table=tblTSv6G3XBYEEF7&view=vewEjey2Cn`

- **App Token:** `Koj9b1Z31aDl9ysTU1dcHus5n2c` ✅
- **Table ID:** `tblTSv6G3XBYEEF7` ✅  
- **View ID:** `vewEjey2Cn` ✅

## Benefits of Using Base vs Sheets

### Lark Base Advantages:
- ✅ Better data validation
- ✅ Structured field types (dates, emails, etc.)
- ✅ Relationships between records
- ✅ Better performance
- ✅ Real-time updates

### When to Use Sheets:
- ✅ Quick setup and testing
- ✅ Importing from Excel
- ✅ Simple data entry
- ✅ No permission setup required