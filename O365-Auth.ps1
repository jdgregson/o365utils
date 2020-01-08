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
        [switch]$Exchange,        # Exchange Online
        [switch]$Azure,           # Azure AD
        [switch]$Security,        # Security and Compliance Center
        [switch]$SharePoint,      # SharePoint Online
        [switch]$SharePointAdmin, # SharePoint Online Administration
        [string]$SharePointURL,
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
            return
        }
    } else {
        $o365creds = Get-Credential
    }

    while ($authAttempts -lt 3 -and -not($authSuccess)) {
        $authAttempts += 1
        if ($Exchange) {
            # Connect to Exchange Online PowerShell
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

        if ($Azure) {
            # Connect to Azure AD
            if (Get-Command "Connect-AzureAD" -errorAction SilentlyContinue) {
                if ($prompt) {
                    $out = Connect-AzureAD | Out-String
                } else {
                    Write-Warning "Warning: You cannot log into AzureAD using an app password. If you are using an app password, please try this command again with the -prompt switch."
                    $out = Connect-AzureAD -Credential $o365creds | Out-String
                }
            } else {
                Write-Warning "Cannot find AzureAD Module. Please run this as an administrator: Install-Module AzureAD"
            }
        }

        if ($Security) {
            # Connect to Security and Compliance Center PowerShell
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

        if ($SharePoint) {
            # Connect to SharePoint Online web and establish a context
            $DLL1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
            $DLL2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
            $libraries = $True
            if (-Not($DLL1) -Or -Not($DLL2)) {
                Write-Host "SharePoint: Could not load required libraries."
                Write-Host "You may need to install the following Windows updates to use this script:"
                Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=35585"
                if ($PSVersionTable -And $PSVersionTable.PSVersion.Major -Lt 5) {
                    Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=50395"
                }
                $libraries = $False
            }

            if ($libraries -and -Not($global:SPcontext) -Or -Not($global:SPweb) -Or -Not($global:SPsite)) {
                if (-not($SharePointURL)) {
                    $SharePointGuess = ($o365creds[0] -split "@")[1]
                    $SharePointURL = "https://$SharePointGuess.sharepoint.com"
                    Write-Warning "No SharePoint URL was given -- assuming `"$SharePointURL`""
                }
                $global:SPurl = $SharePointURL
                $global:SPcontext = New-Object Microsoft.SharePoint.Client.ClientContext($global:SPurl)
                $global:SPcontext.Credentials = $o365creds
                $global:SPcontext.RequestTimeOut = 5000 * 60 * 10;
                $global:SPweb = $global:SPcontext.web
                $global:SPsite = $global:SPcontext.site
                $global:SPcontext.Load($global:SPweb)
                $global:SPcontext.Load($global:SPsite)
                $global:SPcontext.ExecuteQuery()
            }
        }

        if ($SharePointAdmin) {
            # Connect to SharePoint Online Administration
            $DLL1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
            $DLL2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
            $libraries = $True
            if (-Not($DLL1) -Or -Not($DLL2)) {
                Write-Host "SharePointAdmin: Could not load required libraries."
                Write-Host "You may need to install the following Windows updates to use this script:"
                Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=35585"
                if ($PSVersionTable -And $PSVersionTable.PSVersion.Major -Lt 5) {
                    Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=50395"
                }
                $libraries = $False
            }

            if ($libraries) {
                if (-not($SharePointURL)) {
                    $SharePointGuess = (($o365creds.UserName -split "@")[1] -split ".com")[0]
                    $SharePointURL = "https://$SharePointGuess-admin.sharepoint.com"
                    Write-Warning "No SharePoint URL was given -- assuming `"$SharePointURL`""
                } elseif ($SharePointURL -notmatch "-admin.sharepoint.com") {
                    $SharePointURL = $SharePointURL -replace ".sharepoint.com","-admin.sharepoint.com"
                }
                Connect-SPOService -Url $SharePointURL -Credential $o365creds
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
