# o365utils
This is a collection of PowerShell scripts which augment and simplify the
already-powerful PowerShell capabilities of Office 365.

#### Multi-Factor Authentication
Due to issues with how Multi-Factor Authentication has been implemented in the
new Exchange Online Remote PowerShell Module
[discussed here](https://github.com/jdgregson/o365utils/issues/5)
these scripts do not support 2FA directly. As a workaround, you can use an app
password to log in when using these scripts and you will not be prompted for
your 2FA codes. You can use `Generate-CredFiles` to save a copy of the username
and app password to speed up all of these scripts.

#### Requirements
- These scripts were developed in PowerShell v5. If you are using Windows 7 or
  Windows 8.1, please make sure you have installed
  [this](https://www.microsoft.com/en-us/download/details.aspx?id=50395) Windows
  update.
- These scripts were developed using a full Office 365 admin account. If you
have limited privileges in O365, you may have difficulty using some of these
scripts.

#
### Delete-Emails.ps1
Delete-Emails is a PowerShell script used to automate the process of deleting
emails from multiple Office 365 mailboxes at once. Useful for removing malicious
or phishing emails, as well as a more effective method of recalling emails.

Delete-Emails uses the steps discussed
[here](https://support.office.com/en-us/article/3526fd06-b45f-445b-aed4-5ebd37b3762a)
to connect to the Office 365 Security & Compliance Center, perform a search
using your criteria, and delete the emails that it finds (after confirmation, of
course).

```diff
-/!\ WARNING
-Delete Emails is very powerful, likely more powerful than it needs to be. If
-you were to search for and delete emails that were sent to
-brandon@yourcompany.com with no further search limits (such as date and
-subject), this would delete every email that has ever been sent to Brandon by
-any user, inside or outside of your company, including emails in other users'
-inboxes and sent folders (including replys and CC's), junk folders, or deleted
-items folders. This script has the power to remove every single email that has
-ever been sent or received in your Office 365 tenant.
```

**NOTE: If you run the same search again after deleting the results,
Delete-Emails will find the same number of emails again.** Currently this script
performs a SoftDelete on the emails it locates. This means that the emails are
deleted from the users' mailbox and Deleted Items folder, but the user is still
able to use the Recover Deleted Items feature to get the emails back. Running a
search again will show the same number of items, but the script will look at
each mailbox after it is finished and tell you how many of the emails are
_actually_ in the users Inbox.

#
### Get-Forwarders.ps1
Get-Forwarders will produce a list of users who have email forwarding enabled
for their mailbox.

#
### SPO-Lists.ps1
SPO-Lists provides utilities to view, update, and export lists and list items
on SharePoint.


#
### O365-Auth.ps1
O365-Auth is used by these scripts to connect to various Office 365 services
using PowerShell.


#
### Generate-CredFiles.ps1
Generate-CredFiles is used to save usernames and passwords in a credential file.
O365-Auth can use these files to log into Office 365 automatically, bypassing
2FA requirements (if you provide an app password).
