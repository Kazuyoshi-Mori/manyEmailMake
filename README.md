# manyEmailMake
Create multiple emails automatically with this Excel VBA tool!

This project is an Excel macro tool that can automatically generate multiple emails in the Outlook drafts folder. Based on the numerous email addresses in the sheet, the following two methods can be used to create email drafts.

🔧 Main features:

1. BCC-based email creation
   - Create a draft with the multiple email addresses listed in the sheet in the BCC field of one email.
   This is useful when you want to send the same content to multiple recipients without them knowing.

2. CC method of creating individual emails
   - Create a draft email for each email address in the sheet and add other email addresses to the CC field.
   - Send individual emails to each recipient.

🛠️ How to use:

1. Open `manyMailMake.xlsm`.
2. Enter the email addresses in the specified columns of the sheet, and enter the email title and email body in the specified cells.
3. Press the Run button to generate a draft email for each mode automatically in Outlook.

📎 Required environment:

- Excel (with macros enabled in .xlsm format)
- Outlook (must be installed locally)
- VBA-enabled environment
