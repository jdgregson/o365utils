# File: Delete-Emails.ps1
# Purpose: this script will connect to the Office 365 Security & Compliance
# Center and Exchange Online and will allow administrators to delete specified
# emails from every mailbox in their tenant.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>

Param (
    [int]$timeout = "120",
    [switch]$prompt
)

# check if we are on PowerShell version 5 and warn the user if not
if ($PSVersionTable.PSVersion.Major -lt 5) {
    $warning = @"
    ================================ /!\ ================================
    WARNING: Your version of PowerShell is less than V5. This script may
    not run properly in your version. If you run into issues, please
    install this Windows update to bring your PowerShell version to V5:
    https://www.microsoft.com/en-us/download/details.aspx?id=50395
    =====================================================================
"@
    Write-Host $warning -ForegroundColor Yellow
}


function ColorMatch {
    #https://stackoverflow.com/questions/12609760
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $InputObject,
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Pattern,
        [Parameter(Mandatory = $false, Position = 1)]
        [string] $Color='Red'
    )

    begin {$r = [regex]$Pattern}
    process {
        $ms = $r.matches($inputObject)
        $startIndex = 0
        foreach ($m in $ms) {
            $nonMatchLength = $m.Index - $startIndex
            Write-Host $inputObject.Substring($startIndex, $nonMatchLength) -NoNew
            Write-Host $m.Value -Fore $Color -NoNew
            $startIndex = $m.Index + $m.Length
        }
        if ($startIndex -lt $inputObject.Length) {
            Write-Host $inputObject.Substring($startIndex) -NoNew
        }
        Write-Host
    }
}


function New-GUID {
    $guid = [guid]::NewGuid().Guid.Replace("-", "").Substring(25)
    return "delete-emails-$guid"
}


function Delete-Search {
    Param (
        [string]$guid
    )

    Remove-ComplianceSearch -Identity "$guid" -Confirm:$false
}


function Clean-Exit {
    Param (
        [string]$message
    )

    if ($message) {
        Write-Host $message
    }
    Delete-Search "$guid"
    Exit
}


function Get-ComplianceSearchResults {
    Param (
        [string]$guid
    )

    $results = (Get-ComplianceSearch $guid).SuccessResults
    $results = $results -replace "{" -replace "}" -replace "`r`n"
    $results = $results -replace "(, Total size: [0-9,]*)","`r`n"
    $results = $results -split "`r`n"
    return $results
}


function Get-ComplianceSearchResultsUsers {
    Param (
        [string]$guid
    )

    $results = Get-ComplianceSearchResults $guid
    $usersWithResults = @()
    $pattern = "Location: (.*?), Item count: [0-9]?"
    foreach ($mailbox in $results) {
        if ([int]($mailbox.Split(' ')[4]) -gt 0) {
            $usersWithResults += [regex]::match($mailbox, $pattern).Groups[1].Value
        }
    }
    return $usersWithResults
}


function Get-ComplianceSearchResultsList {
    Param (
        [string]$guid
    )

    $results = Get-ComplianceSearchResults $guid
    foreach ($mailbox in $results) {
        if ([int]($mailbox.Split(' ')[4]) -gt 0) {
            "$mailbox" | ColorMatch "Item count: [0-9]*"
        }
    }
}


function Get-ComplianceSearchResultsPreview {
    Param (
        [string]$guid
    )

    New-ComplianceSearchAction -SearchName $guid -Preview -ErrorAction SilentlyContinue
    $name = "$($guid)_Preview"
    $preview = Get-ComplianceSearchAction $name
    Write-Host "Creating preview..."
    if ($preview -ne $Null -and $preview.GetType().Name -eq "PSObject") {
        while ($preview.Status -ne "Completed") {
            sleep 1
            $preview = Get-ComplianceSearchAction $name
        }
        Write-Host "Formatting preview..."
        $outputTable = @()
        $items = $preview.Results -replace "{" -replace "}" -split ".eml,"
        foreach ($_i in $items) {
            $item = $_i -split ";" -replace "^ " -replace "`n"
            $i = New-Object -Type PSObject -Property @{
                To = $item[0] -Replace "Location: "
                From = $item[1] -Replace "Sender: "
                Subject = $item[2] -Replace "Subject: "
                Received = $item[5] -Replace "Received Time: "
            }
            $outputTable += $i
        }
        Get-ComplianceSearchAction $name | Remove-ComplianceSearchAction -Confirm:$False
        return $outputTable
    } else {
        Write-Warning "Could not preview results. Are you a member of the eDiscovery Managers or eDiscovery Administrators group?"
        Write-Warning "See here for details: https://docs.microsoft.com/en-us/office365/securitycompliance/assign-ediscovery-permissions"
        return $Null
    }
}


function Test-ComplianceSearchComplete {
    Param (
        [string]$guid
    )

    $theSearch = Get-ComplianceSearch -Identity "$guid" | Format-List -Property Status | Out-String
    $searchProgress = $theSearch | Select-String -pattern "Completed"
    if ($searchProgress.length -gt 0) {
        return $true
    } else {
        return $false
    }
}


# set up our remote Office 365 PowerShell sessions
. "$PSScriptRoot\O365-Auth.ps1"
if (($prompt -and (O365-Auth -Exchange -Security -Prompt) -eq 1) -or (O365-Auth -Exchange -Security) -eq 1) {
    Write-Warning "Failed to authenticate with Office 365"
    Exit
}

$examples =
'Example: sent>=07/03/2017 AND sent<=07/05/2017 AND subject:"open this attachment!"',
'Example: subject:"contains this phrase" from:somedomain.com',
'Example: to:user@mycompany.com',
'Example: from:some.spammer@hijackeddomain.com',
'Example: attachment:"Malicious-File.docx"',
'Example: attachment:"docx" NOT from:user@mycompany.com',
'More: https://technet.microsoft.com/en-us/library/ms.exch.eac.searchquerylearnmore(v=exchg.150).aspx'

# get the search criteria from the user
while ($true) {
    if ($search -and $search.ToUpper() -eq 'M') {
        $examples
    } elseif ($search) {
        $search = "kind:email $search"
        break
    } else {
        Write-Host "Enter a search string to locate the email(s)"
        Write-Host $examples[0]
    }
    $search = Read-Host "(enter `"M`" for more examples) Search"
}

# create and run the search
$guid = New-GUID
$out = New-ComplianceSearch -Name $guid -ExchangeLocation all -ContentMatchQuery "$search" | Out-String
Write-Host "Starting the search..."
Start-ComplianceSearch $guid

# wait for the results and ask the user if they look right
$searchCompleted = $false
$usersWithResults = @()
for ($i=0; $i -le $timeout; $i++) {
    if (Test-ComplianceSearchComplete($guid)) {
        $searchCompleted = $true
        Write-Host "Search complete"
        Write-Host "The search returned the following:"
        Get-ComplianceSearch $guid | Format-List -Property Items
        if ((Get-ComplianceSearch $guid).Items -eq 0) {
            Clean-Exit "0 items found. Cleaning up and exiting..."
        }
        $usersWithResults = Get-ComplianceSearchResultsUsers $guid
        Write-Host "Does this seem accurate?"
        $answer = Read-Host "[Y] Yes  [N] No  [M] More details [P] Preview results (default is `"N`")"
        if ($answer.ToUpper() -eq "Y") {
            Write-Host "Confirmed. Continuing to delete..."
            break
        } elseif ($answer.ToUpper() -eq "M") {
            Get-ComplianceSearchResultsList $guid
            continue;
        } elseif ($answer.ToUpper() -eq "P") {
            $previewTable = (Get-ComplianceSearchResultsPreview $guid) | Select-Object -skip 1
            if ($previewTable -eq $Null) {
                continue;
            }
            $previewTable | Sort-Object To | Format-Table -AutoSize -Wrap
            while ($True) {
                $answer = Read-Host "`n[C] Continue  [L] List output  [T] Table output"
                if ($answer.ToUpper() -eq "C") {
                    break;
                } elseif ($answer.ToUpper() -eq "L") {
                    $previewTable | Sort-Object To | Format-List
                } elseif ($answer.ToUpper() -eq "T") {
                    $previewTable | Sort-Object To | Format-Table -AutoSize -Wrap
                }
            }
            continue;
        } else {
            Clean-Exit "Canceled. Cleaning up and exiting..."
        }
    }
    Sleep 1
}

if ($searchCompleted -eq $false) {
    "Error: the search timed out" | ColorMatch .
    "Try running this script with a longer timeout, e.g:" | ColorMatch .
    "    Delete-Emails -Timeout 6000" | ColorMatch .
    Clean-Exit
}

# delete the emails with the user's confirmation
$out = New-ComplianceSearchAction -SearchName "$guid" -Purge -PurgeType SoftDelete | Out-String
$ComplianceSearchActions = Get-ComplianceSearchAction | Out-String
$purgeProgress = $ComplianceSearchActions | Select-String -Pattern $guid
# if the user did not confirm then exit
if ($purgeProgress.length -eq 0) {
    Clean-Exit "The purge was canceled. Cleaning up and exiting..."
}

# wait for the deletion results and delete the search if it is finished
for ($i=0; $i -le $timeout; $i++) {
    $thePurge = Get-ComplianceSearchAction -Identity $guid"_Purge" | Out-String
    $purgeProgress = $thePurge | Select-String -Pattern "Completed"
    if ($purgeProgress.length -gt 0) {
        Write-Host "Deletion complete"
        Delete-Search "$guid"
        Break
    }
    Sleep 1
}

$confMessage = "Would you like to confirm the deletion? This will start many searches and may take a while."
$confMessage = "$confMessage`n[Y] Yes  [N] No  (default is `"N`")"
if (-not((Read-Host $confMessage).ToUpper() -eq "Y")) {
    Write-Host "Skipping confirmation"
    Clean-Exit
}
# for each mailbox with results, create a search query which will exclude
# deleted items folders
# see: https://support.office.com/en-us/article/e3cbc79c-5e97-43d3-8371-9fbc398cd92e
Write-Host "Confirming deletion..."
$PendingDeletions = New-Object System.Collections.ArrayList(,@($usersWithResults))
$ConfirmationSearches = New-Object System.Collections.ArrayList
for ($i=0; $i -lt $PendingDeletions.Count; $i++) {
    $UserEmail = $PendingDeletions[$i]
    $folderExclusionsQuery = " AND NOT ("
    $excludeFolders = "/Deletions","/Purges","/Recoverable Items"
    $folderStatistics = Get-MailboxFolderStatistics $UserEmail
    foreach ($folderStatistic in $folderStatistics) {
        $folderPath = $folderStatistic.FolderPath;
        if ($excludeFolders.Contains($folderPath)) {
            $folderId = $folderStatistic.FolderId;
            $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
            $nibbler= $encoding.GetBytes("0123456789ABCDEF");
            $folderIdBytes = [Convert]::FromBase64String($folderId);
            $indexIdBytes = New-Object byte[] 48;
            $indexIdIdx=0;
            $folderIdBytes | select -skip 23 -First 24 | %{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
            $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
            $folderExclusionsQuery += "($folderQuery) OR "
        }
    }
    $folderExclusionsQuery += ")"
    $fullSearch = "$UserEmail#$search $folderExclusionsQuery"
    $PendingDeletions[$i] = $fullSearch
}

$MailboxesWithResults = $PendingDeletions.Count
while ($PendingDeletions.Count -gt 0) {
    foreach ($PendingDeletion in $PendingDeletions) {
        $PendingDeletion = $PendingDeletion -Split '#'
        $thisGuid = New-GUID
        $out = New-ComplianceSearch -Name "$thisGuid" -ExchangeLocation $PendingDeletion[0] -ContentMatchQuery "$($PendingDeletion[1])" | Out-String
        Start-ComplianceSearch -Identity "$thisGuid"
        [void]$ConfirmationSearches.Add($thisGuid)
    }
    while ($ConfirmationSearches.Count -gt 0) {
        for ($i=0; $i -lt $ConfirmationSearches.Count; $i++) {
            $thisSearch = $ConfirmationSearches[$i];
            if (Test-ComplianceSearchComplete("$thisSearch")) {
                $results = Get-ComplianceSearchResults "$thisSearch";
                $thisQuery = (Get-ComplianceSearch $thisSearch).ContentMatchQuery
                $thisUser = (Get-ComplianceSearch $thisSearch).ExchangeLocation
                $ConfirmationSearches.Remove($thisSearch)
                Delete-Search "$thisSearch"
                foreach ($mailbox in $results) {
                    if ($mailbox -and [int]($mailbox.Split(' ')[4]) -eq 0) {
                        $PendingDeletions.Remove("$thisUser#$thisQuery")
                        $Progress = "($($MailboxesWithResults-$PendingDeletions.Count)/$MailboxesWithResults)"
                        "$Progress $mailbox" -replace('Location: ') | ColorMatch "Item count: [0-9]*" -Color 'Green'
                    }
                }
            }
        }
        sleep 0.5
    }
}
