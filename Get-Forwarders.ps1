# File: Get-Forwarders.ps1
# Purpose: this script will connect to the Office 365 Exchange Online PowerShell
# and produce a list of user who are forwarding their email to another address.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>


Param (
    [switch]$ShowAll = $False,
    [string]$CredFile,
    [switch]$Prompt
)


# connect to Exchange Online PowerShell
. "$PSScriptRoot\O365-Auth.ps1"
if (($prompt -and (O365-Auth -Exchange -Prompt) -eq 1) -or (O365-Auth -Exchange) -eq 1) {
    Write-Warning "Failed to authenticate with Office 365"
    Exit
}


$mailboxes = (Get-Mailbox | Select-Object UserPrincipalName,ForwardingSmtpAddress,DeliverToMailboxAndForward)
$mailboxes = $mailboxes | Where-Object {$_.UserPrincipalName -notmatch "DiscoverySearchMailbox"}
if ($ShowAll) {
    $mailboxes
} else {
    $mailboxes | Where-Object {$_.ForwardingSmtpAddress -ne $Null}
}
