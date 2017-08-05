# File: Delete-Emails.ps1
# Purpose: this script will connect to the Office 365 Security & Compliance
# Center and will allow admins to delete specified emails from every mailbox
# in the organization.
# Author: Jonathan Gregson <jonathan.gregson@agsstainless.com>
#                          <jdgregson@gmail.com>

# check if we are on PowerShell version 5 and warn the user if not
$psversion = $PSVersionTable.PSVersion | Format-List -Property Major | Out-String
$psversion = [int]($psversion -split ": ")[1]
if($psversion -lt 5) {
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

# connect to Office 365 Security & Compliance Center
try{$out = Get-ComplianceSearch|Out-String} catch {
    Write-Host "Enter your Office 365 admin user credentials..."
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber -DisableNameChecking
    $Host.UI.RawUI.WindowTitle = $UserCredential.UserName + " (Office 365 Security & Compliance Center)"
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
while($true) {
    if($search -And $search -eq 'm' -Or $search -eq 'M') {
        $examples
    } elseif($search) {
        $search = "kind:email $search"
        break
    } else {
        Write-Host "Enter a search string to locate the email(s)"
        Write-Host $examples[0]
    }
    $search = Read-Host "(enter 'm' for more examples) Search"
}

# create and run the search
$guid = [guid]::NewGuid().Guid.Replace("-", "").Substring(25)
$out = New-ComplianceSearch -Name "$guid" -ExchangeLocation all -ContentMatchQuery "$search" | Out-String
Write-Host "Starting the search..."
Start-ComplianceSearch -Identity "$guid"

# wait for the results and ask the user if they look right
$searchCompleted = $false
$timeout = 60
For ($i=0; $i -le $timeout; $i++) {
    $theSearch = Get-ComplianceSearch -Identity "$guid" | Format-List -Property Status | Out-String
    $searchProgress = $theSearch | Select-String -pattern "Completed"
    if($searchProgress.length -gt 0) {
        $searchCompleted = $true
        Write-Host "Search complete"
        Write-Host "The search returned the following:"
        Get-ComplianceSearch -Identity "$guid" | Format-List -Property Items
        Write-Host "Does this seem accurate?"
        Write-Host "[Y] Yes [N] No [M] More details - default No"
        $answer = Read-Host "Confirm"
        if($answer.ToLower() -eq "y") {
            Write-Host "Confirmed. Continuing to delete..."
            break
        } elseif($answer.ToLower() -eq "m") {
            # if the user asked for more details, parse the results and show
            # only the mailboxes that have items which were found
            $results = Get-ComplianceSearch -Identity "$guid" | Format-List -Property SuccessResults | Out-String
            $results = $results -Replace "SuccessResults : {"
            $results = $results -Replace "                 "
            $results = $results -Replace "}"
            $results = $results -Replace "`r`n"
            $results = $results -Replace "(Total size: [0-9,]*)","`r`n"
            $results = $results -split "`r`n"
            ForEach($mailbox in $results) {
                if([int]($mailbox.Split(' ')[4]) -gt 0) {
                    "$mailbox" | ColorMatch "Item count: [0-9]*"
                }
            }
            continue;
        } else {
            Write-Host "Canceled. Cleaning up and exiting..."
            Remove-ComplianceSearch -Identity "$guid" -Confirm:$false
            Exit
        }
    }
    Sleep 1
}
if($searchCompleted -eq $false) {
    Write-Host "Error: the search timed out"
    Remove-ComplianceSearch -Identity "$guid" -Confirm:$false
    Exit
}

# delete the emails with the user's confirmation
$out = New-ComplianceSearchAction -SearchName "$guid" -Purge -PurgeType SoftDelete | Out-String

# wait for the deletion results and delete the search if it is
$timeout = 120
For ($i=0; $i -le $timeout; $i++) {
    $thePurge = Get-ComplianceSearchAction -Identity $guid"_Purge" | Out-String
    $purgeProgress = $thePurge | Select-String -Pattern "Completed"
    if($purgeProgress.length -gt 0) {
        Write-Host "Deletion complete"
        Write-Host "Cleaning up and exiting..."
        Remove-ComplianceSearch -Identity "$guid" -Confirm:$false
        Exit
    }
    Sleep 1
}

function ColorMatch {
    #https://stackoverflow.com/questions/12609760
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $InputObject,

        [Parameter(Mandatory = $true, Position = 0)]
        [string] $Pattern
    )
    begin{ $r = [regex]$Pattern }
    process {
        $ms = $r.Matches($inputObject)
        $startIndex = 0
        foreach($m in $ms) {
            $nonMatchLength = $m.Index - $startIndex
            Write-Host $inputObject.Substring($startIndex, $nonMatchLength) -NoNew
            Write-Host $m.Value -Fore DarkRed -NoNew
            $startIndex = $m.Index + $m.Length
        }
        if($startIndex -lt $inputObject.Length) {
            Write-Host $inputObject.Substring($startIndex) -NoNew
        }
        Write-Host
    }
}