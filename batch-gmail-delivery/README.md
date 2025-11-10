# Batch Gmail Delivery

A Google Apps Script that performs mail merge, sending personalized emails from Gmail using a Google Doc template and Google Sheet data source.

Google Sheet Example https://docs.google.com/spreadsheets/d/1_ffSSnm2Bail9Aulk2Yr0X_MXt9mQTXDPDQB1_Z3rWc/edit?gid=0#gid=0

## Features

- Send bulk personalized emails using Gmail
- Use Google Doc as email template with variable placeholders
- Control sending with status column (Ready/WIP)
- Track sent emails with checkbox column
- All user-facing messages in Traditional Chinese (zh-TW)

## Setup

1. Copy the code from `Code.ts` into Google Apps Script (save as `Code.gs`)
2. In your Google Sheet, go to **批量寄送Gmail > 建立/重置設定表單**
3. Fill in the Settings sheet:
   - **Google Doc Template ID**: The ID of your Google Doc template
   - **Data Sheet Name**: Name of the sheet containing your data (default: "客戶資料")
   - **Email Column Name**: Name of the column containing email addresses (default: "客戶信箱")

## Sheet Format

Your data sheet must include these **required columns**:

| Column Name | Type | Description |
|------------|------|-------------|
| `客戶信箱` (or your configured email column) | Text | Recipient email address |
| `已寄送？` | Checkbox | Automatically checked when email is sent |
| `狀態` | Text | Must be "Ready" to send emails. Use "WIP" to prevent sending |

### Additional Columns

You can add any other columns you want. Their values can be used in the email template using `{{Column Name}}` placeholders.

## Email Template Format

Create a Google Doc with:
- **First line**: Email subject (can use `{{Column Name}}` placeholders)
- **Remaining lines**: Email body (can use `{{Column Name}}` placeholders)

Example:
```
Hello {{Name}}, your invoice is ready
Dear {{Name}},

Your invoice for {{Amount}} is ready. Please review and let us know if you have any questions.

Best regards,
{{Company Name}}
```

## How It Works

1. The script reads rows from your data sheet
2. For each row, it checks:
   - `已寄送？` checkbox is **unchecked**
   - `狀態` is **"Ready"**
   - Email address is **not empty**
3. If all conditions are met, it:
   - Replaces `{{Column Name}}` placeholders with actual values
   - Sends the email via Gmail
   - Checks the `已寄送？` checkbox

## Usage

1. Prepare your data sheet with the required columns
2. Set rows you want to send to `狀態 = "Ready"`
3. Set rows you want to skip to `狀態 = "WIP"`
4. Go to **批量寄送Gmail > 寄送Gmail**
5. The script will send emails and update the checkbox column automatically

