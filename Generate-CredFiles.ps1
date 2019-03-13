# File: Generate-CredFile.ps1
# Purpose: prompts for a username and password and stores them as secure strings
# in the specified file.
#
# Copyright (c) 2018, AGS Stainless Inc.
# Author: Jonathan Gregson <Jonathan.Gregson@agsstainless.com>
#                          <jdgregson@gmail.com>

Param (
    [string]$OutputFile = "$PSScriptRoot\default.creds"
)

$username = Read-Host "Enter the username to store"
$password = Read-Host -AsSecureString "Enter the password to store"
$combined = "$username##$(ConvertFrom-SecureString $password)"
$combined > $OutputFile
Write-Host "Wrote credentials to `"$OutputFile`""
