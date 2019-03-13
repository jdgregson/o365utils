# File: O365-Auth.ps1
# Purpose: this file handles remote PowerShell sessions with the Office 365
# backend and simplifies the process of logging into various Office 365
# services.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>


function Read-CredFile {
    Param (
        [string]$CredFile
    )

    $combined = Get-Content $CredFile
    $username = ($combined -split "##")[0]
    $password = ($combined -split "##")[1] | ConvertTo-SecureString
    return New-Object System.Management.Automation.PSCredential ($username, $password)
}


function O365-Auth {
    Param (
        [switch]$Exchange,  # Exchange Online
        [switch]$Security,  # Security and Compliance Center
        [string]$CredFile,
        [switch]$Prompt
    )
    $authAttempts = 0
    $authSuccess = $false
    $o365creds = $Null

    if (($CredFile -or $(Test-Path "$PSScriptRoot\default.creds")) -and -not($prompt)) {
        if (-not($CredFile)) {
            $CredFile = "$PSScriptRoot\default.creds"
        }
        if (Test-Path $CredFile) {
            $o365creds = (Read-CredFile -CredFile $CredFile)
        } else {
            Write-Warning "`"$CredFile`" is not a valid credential file."
            Exit
        }
    } else {
        $o365creds = Get-Credential
    }

    while ($authAttempts -lt 3 -and -not($authSuccess)) {
        $authAttempts += 1
        if ($Exchange) {
            # connect to Exchange Online PowerShell
            if (-not($global:EOSession) -or($global:EOSession.State -ne "Opened") -or($global:EOSession.Availability -ne "Available")) {
                try {
                    $global:EOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid/ -Credential $o365creds -Authentication Basic -AllowRedirection
                    $out = Import-PSSession $global:EOSession -AllowClobber -DisableNameChecking|Out-String
                    if ($out -like "*ExportedCommands*") {Write-Host "Successfully connected to Exchange Online"}
                } catch {
                    $o365creds = $Null
                    Write-Warning "Exchange Online login failed, please try again"
                    Continue
                }
            }
        }

        if ($Security) {
            # connect to Security and Compliance Center PowerShell
            if (-not($global:SCCSession) -or($global:SCCSession.State -ne "Opened") -or($global:SCCSession.Availability -ne "Available")) {
                try {
                    $global:SCCSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $o365creds -Authentication Basic -AllowRedirection
                    $out = Import-PSSession $global:SCCSession -AllowClobber -DisableNameChecking|Out-String
                    if ($out -like "*ExportedCommands*") {Write-Host "Successfully connected to Security and Compliance Center"}
                } catch {
                    $o365creds = $Null
                    Write-Warning "Security and Compliance Center login failed, please try again"
                    Continue
                }
            }
        }

        $authSuccess = $true
        $o365creds = $Null
    }

    if ($authSuccess) {
        return 0
    } else {
        return 1
    }
}