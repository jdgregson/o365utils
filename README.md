# Delete Emails (O365)
Delete-Emails is a PowerShell script used to automate the process of deleting
emails from multiple Office 365 mailboxes at once. Useful for removing malicious
or phishing emails, as well as a more effective method of recalling emails.

## What it Does
Delete-Emails uses the steps discussed [here](https://support.office.com/en-us/article/3526fd06-b45f-445b-aed4-5ebd37b3762a)
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

## Requirements
- Delete-Emails uses some features in PowerShell v5. If you are using Windows 7
or Windows 8.1, please make sure you have installed [this](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
  Windows update.
- This script was developed using a full Office 365 admin account. If you have
limited privileges in O365, you may have difficulty using this script.

## FAQ
**I ran the search again after they were deleted and found the same number of
emails. What gives?**
Currently this script performs a SoftDelete on the emails it locates. This means
that the emails are deleted from the users' mailbox and Deleted Items folder,
but the user is still able to use the Recover Deleted Items feature to get the
emails back. Future versions of this script may provide a way to permanently
remove the items, or at least tell you how many of the results are already
deleted.

## TODO
- Add a way to show how many of the returned results are in the Recoverable
Items box so that admins don't think the emails are still there.