# File: End-O365Sessions.ps1
# Purpose: this script will use all available methods to end all sessions of a
# specific user as soon as possible. Unfortunately, it may take up to an hour
# for some sessions on some devices to be effectively terminated.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>

Param (
    [string]$User,
    [int]$Attempts = 100,
    [string]$CredFile,
    [switch]$Prompt
)

# connect to Azure AD and SharePoint Online Administration
. "$PSScriptRoot\O365-Auth.ps1"
if (($prompt -and (O365-Auth -Azure -SharePointAdmin -Prompt) -eq 1) -or (O365-Auth -Azure -SharePointAdmin) -eq 1) {
    Write-Warning "Failed to authenticate with Office 365"
    Exit
}

Write-Host "Attempting to end all sessions $Attempts times..."
while ($Attempts -gt 0) {
    Revoke-SPOUserSession -User $User -Confirm:$False -WarningAction "SilentlyContinue"
    Get-AzureAdUser -SearchString $User | Revoke-AzureADUserAllRefreshToken
    $Attempts--
}
