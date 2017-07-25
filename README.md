# Delete Emails (O365)
Delete-Emails is a PowerShell script used to automate the process of deleting
emails from multiple Office 365 mailboxes at once. Useful for removing malicious
or phishing emails, as well as a more effective method of recalling emails.

## What it Does
Delete-Emails uses the steps discussed [here](https://support.office.com/en-us/article/3526fd06-b45f-445b-aed4-5ebd37b3762a)
to connect to the Office 365 Security & Compliance Center, perform a search 
using your criteria, and delete the emails that it finds (after confirmation, of 
course).

`````/!\ WARNING`````
`````Delete Emails is very powerful, likely more powerful than it needs to `````
`````be. If you were to search for and delete emails that were sent to `````
`````brandon@yourcompany.com with no further search limits (such as date `````
`````and subject), this would delete every email that has ever been sent to`````
````` Brandon by any user, inside or outside of your company, including `````
`````emails in other users' inboxes and sent folders (including replys and`````
````` CC's), junk folders, or deleted items folders. This script has the `````
`````power to remove every single email that has ever been sent in your `````
`````Office 365 tenant. `````

## Requirements
- Delete-Emails uses some featues in PowerShell v5. If you are using Windows 7
or Windows 8.1, please make sure you have installed [this](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
  Windows update.
- This script was developed using a full Office 365 admin account. If you have
limited privelages in O365, you may have difficulty using this script.