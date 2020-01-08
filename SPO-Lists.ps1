# File: SPO-Lists.ps1
# Purpose: provides some utilities for working with list in a SharePoint Online
# tenant.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>

Param (
    [bool]$import = $False,
    [bool]$silent = $False
)


function SPOLists-Init {
    Param (
        [string]$sharePointUrl,
        [string]$username,
        [Security.SecureString]$password,
        [bool]$silent = $False
    )

    $DLL1 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    $DLL2 = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    if (-Not($DLL1) -Or -Not($DLL2)) {
        Write-Host "Could not load shared libraries."
        Write-Host "You may need to install the following Windows updates to use this script:"
        Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=35585"
        if ($PSVersionTable -And $PSVersionTable.PSVersion.Major -Lt 5) {
            Write-Host "    https://www.microsoft.com/en-us/download/details.aspx?id=50395"
        }
        Exit
    }

    if (-Not($global:SPcontext) -Or -Not($global:SPweb) -Or -Not($global:SPsite)) {
        $global:SharePointUrl = $sharePointUrl
        Write-Host "Connecting to $global:sharePointUrl..."
        if ($username) {
            Write-Host "Username: $username"
        } else {
            $username = Read-Host "Enter username"
        }
        if (-Not($password)) {
            $password = Read-Host -Prompt "Enter password" -AsSecureString
        }
        $global:SPusername = $username
        $global:SPcontext = New-Object Microsoft.SharePoint.Client.ClientContext($global:SharePointUrl)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($global:SPusername, $password)
        $global:SPcontext.Credentials = $credentials
        $global:SPcontext.RequestTimeOut = 5000 * 60 * 10;
        $global:SPweb = $global:SPcontext.web
        $global:SPsite = $global:SPcontext.site
        $global:SPcontext.Load($global:SPweb)
        $global:SPcontext.Load($global:SPsite)
        $global:SPcontext.ExecuteQuery()
    }

    if (-Not($silent)) {
        "Successfully connected to the SharePoint online site $global:SharePointUrl."
        "Some utilities are provided to help work with lists, as detailed below.`n"

        Write-Host "NOTE: if you get an error saying 'The collection has not been initialized.', you likely need to assign the value in question to a varialbe and loop through it." -Foreground Yellow

        show-listcommands
        Write-Host "type Show-ListCommands to show this list again"

        Write-Host "`nType `"Get-Help command -Full`" for more details. I.E. `"Get-Help Get-ListFields -Full`""

    }
}


function global:blue($Output, $NewLine=$true) {
    if ($NewLine) {Write-Host "$Output" -ForegroundColor Cyan}
    else {Write-Host "$Output" -ForegroundColor Cyan -NoNewline}
}


function global:magenta($Output, $NewLine=$true) {
    if ($NewLine) {Write-Host "$Output" -ForegroundColor Magenta}
    else {Write-Host "$Output" -ForegroundColor Magenta -NoNewline}
}


function global:Show-ListCommands {
    Write-Host "`nAvailable Commands:"
    magenta "Write-ValueToListField"  $false
    blue " ListName FieldName ReplaceWith"
    magenta "Write-ValueToListFieldIf" $false
    blue " ListName FieldName IsEqualTo ReplaceWith"
    magenta "Get-ListFields" $false
    blue " ListName FieldNames" $false
    Write-Host " - E.G.: Get-ListFields 'My List' 'Title','Field1','Field2'"
    magenta "Get-AllListItems" $false
    blue " ListName"
    magenta "Show-AllListItems" $false
    blue " ListName"
    magenta "Get-UserListItems" $false
    blue " ListName"
    magenta "Show-UserListItems" $false
    blue " ListName"
    magenta "Get-AllListNames"
    magenta "Show-AllListNames"
    magenta "Get-UserListNames"
    magenta "Show-UserListNames"
    magenta "Get-AllFieldNames"  $false
    blue " ListName"
    magenta "Show-AllFieldNames"  $false
    blue " ListName"
    magenta "Get-UserFieldNames"  $false
    blue " ListName"
    magenta "Show-UserFieldNames"  $false
    blue " ListName"
}

$global:SP_SYSTEM_FIELDS = 'Content Type ID','Approver Comments','File Type',
    'Created By','App Modified By','Last Modified Date','Total File Stream Size',
    'Modified By','Modified','Created','ID','Content Type','Property Bag',
    'Has Copy Destinations','Copy Source','owshiddenversion','Name','Type',
    'Workflow Version','UI Version','Version','Attachments','Approval Status',
    'Edit','Select','Instance ID','Order','GUID','Workflow Instance ID',
    'URL Path','Path','Item Type','Sort Type','Effective Permissions Mask',
    'Unique Id','Client Id','ProgId','ScopeId','HTML File Type','Label setting',
    'Edit Menu Table Start','Edit Menu Table End','Server Relative URL',
    'Encoded Absolute URL','File Name','Level','Is Current Version','Total Size',
    'Item Child Count','Folder Child Count','Restricted','Originator Id',
    'NoExecute','Content Version','Labels','Label Applied','Label applied by',
    'Access Policy','VirusStatus','VirusVendorID','VirusInfo','App Created By',
    'Total File Count','Compliance Asset Id','Item is a Record'

$global:SP_SYSTEM_LISTS = 'Access Requests','appdata','appfiles','Composed Looks',
    'Content type publishing error log','Converted Forms','Form Templates',
    'fpdatasources','List Template Gallery','Maintenance Log Library','wfpub',
    'Master Page Gallery','Project Policy Item List','SharePointHomeOrgLinks',
    'Sharing Links','Solution Gallery','Style Library','TaxonomyHiddenList',
    'Theme Gallery','User Information List','Web Part Gallery','Site Pages',
    'Site Assets'


function global:Write-ValueToListField($ListName, $FieldName, $ReplaceWith, $silent=$false, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Loops through every item in $ListName and will replace the contents of the $field field with $replaceWith.
    .PARAMETER ListName
    The display name of the list to be updated.
    .PARAMETER FieldName
    The name of the field to be updated.
    .PARAMETER ReplaceWith
    The value to insert into the specified field.
    .PARAMETER Silent
    [bool] whether or not to print details of the changes.
    .NOTES
    THIS CANNOT BE EASILY REVERSED.
    .EXAMPLE
    Write-ValueToListField "My List" "Status" "Shipped"
    #>

    $Fields = Get-AllFieldNames $ListName $w $c
    if (-Not($Fields.Contains($FieldName))) {
        Write-Host "'$FieldName' is not a field in '$ListName'. Did you use proper capitalization?" -ForegroundColor Red
        return
    }
    $List = Get-AllListItems $ListName $w $c
    ForEach ($ListItem in $List) {
        if (-Not([string]::IsNullOrEmpty($ReplaceWith)) -And $ReplaceWith.Substring(0, 1) -Eq '>' -And -Not($ReplaceWith.Substring(1, 2) -Eq '>')) {
            $_ReplaceWith = $ReplaceWith.TrimStart('>')
            $_ReplaceWith = Invoke-Expression $_ReplaceWith
        } else {
            $_ReplaceWith = $ReplaceWith
        }
        $OldValue = $ListItem["$FieldName"];
        $ListItem["$FieldName"] = "$_ReplaceWith"
        $ListItem.Update()
        if (-Not($silent)) {
            $Title = $ListItem["Title"]
            $NewValue = $ListItem["$FieldName"];
            if (-Not($NewValue)) {$NewValue = '[EMPTY]'}
            if (-Not($OldValue)) {$OldValue = '[EMPTY]'}
            Write-Host "$Title.$FieldName, $OldValue --> $NewValue"
        }
        $c.Load($ListItem)
        $c.ExecuteQuery()
    }
}


function global:Write-ValueToListFieldif ($ListName, $FieldName, $IsEqualTo, $ReplaceWith, $silent=$false, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Loops through every item in $ListName and will replace the contents of the $field field with $replaceWith.
    .PARAMETER ListName
    The display name of the list to be updated.
    .PARAMETER FieldName
    The name of the field to be updated.
    .PARAMETER IsEqualTo
    The value that $FieldName must be in order to be replaced by $ReplaceWith.
    .PARAMETER ReplaceWith
    The value to insert into the specified field.
    .PARAMETER Silent
    [bool] whether or not to print details of the changes.
    .NOTES
    THIS CANNOT BE EASILY REVERSED.
    .EXAMPLE
    Write-ValueToListFieldIf "My List" "Status" "Processing" "Shipped"
    #>

    $Fields = Get-AllFieldNames $ListName $w $c
    if (-Not($Fields.Contains($FieldName))) {
        Write-Host "'$FieldName' is not a field in '$ListName'. Did you use proper capitalization?" -ForegroundColor Red
        return
    }
    $List = Get-AllListItems $ListName $w $c
    $CleanFieldName = $FieldName.Replace(" ", "_x0020_")
    ForEach ($ListItem in $List) {
        if (-Not([string]::IsNullOrEmpty($ReplaceWith)) -And $ReplaceWith.Substring(0, 1) -Eq '>' -And -Not($ReplaceWith.Substring(1, 2) -Eq '>')) {
            $_ReplaceWith = $ReplaceWith.TrimStart('>')
            $_ReplaceWith = Invoke-Expression $_ReplaceWith
        } else {
            $_ReplaceWith = $ReplaceWith
        }
        $OldValue = $ListItem["$CleanFieldName"];
        if ($OldValue -Eq $IsEqualTo) {
            $ListItem["$CleanFieldName"] = "$_ReplaceWith"
            $ListItem.Update()
            $c.Load($ListItem)
            $c.ExecuteQuery()
        }
        if (-Not($silent)) {
            $Title = $ListItem["Title"][0..25] -Join ""
            $NewValue = $ListItem["$CleanFieldName"];
            if ([string]::IsNullOrEmpty($NewValue)) {$NewValue = '[EMPTY]'}
            if ([string]::IsNullOrEmpty($OldValue)) {$OldValue = '[EMPTY]'}
            if ($OldValue -Eq $IsEqualTo) {
                Write-Host "$Title.$FieldName, $OldValue --> $NewValue" -ForegroundColor Green
            } else {
                Write-Host "$Title.$FieldName, $OldValue --> $NewValue" -ForegroundColor Red
            }
        }
    }
}


function global:Get-AllListItems($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns every row and field in $ListName as an array.
    .PARAMETER ListName
    The display name of the list to be returned.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $Items = Get-AllListItems "My List"
    #>

    $Lists = $w.Lists
    $c.Load($Lists)
    $c.ExecuteQuery()
    $List = $Lists.getByTitle($ListName)
    $c.Load($List)
    $c.ExecuteQuery()
    $Items = $List.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())
    $c.Load($Items)
    $c.ExecuteQuery()
    return $Items
}


function global:Show-AllListItems($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows every row and field in $ListName.
    .PARAMETER ListName
    The display name of the list to be shown.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $Items = Get-AllListItems "My List"
    #>

    $ListItems = Get-AllListItems $ListName $w $c
    $ListFields = Get-AllFieldNames $ListName $w $c
    ForEach ($ListItem in $ListItems) {
        ForEach ($ListField in $ListFields) {
            Write-Host $ListItem["$ListField"]
        }
    }
}


function global:Get-UserListItems($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns the data in every user-created field in each row of $ListName in CSV format.
    .PARAMETER ListName
    The display name of the list to be returned.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $Items = Get-UserListItems "My List"
    #>

    $Fields = Get-UserFieldNames $ListName $w $c
    $Items = Get-ListFields $ListName $Fields $w $c
    return $Items
}


function global:Show-UserListItems($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows the data in every user-created field in each row of $ListName in CSV format.
    .PARAMETER ListName
    The display name of the list to be returned.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $Items = Get-UserListItems "My List"
    #>

    Get-UserListItems $ListName $w $c
}


function global:Get-ListFields($ListName, $FieldNames, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns specified fields from all rows in $ListName in CSV format.
    .PARAMETER ListName
    The display name of the list with fields to be returned.
    .PARAMETER FieldNames
    The display name of the fields to return from the list (as an array). E.G:
    Get-ListFields 'My List' 'Title','Field1','Field2'
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $Items = Get-ListFields 'My List' 'Title','Field1','Field2'
    #>

    $ListItems = Get-AllListItems $ListName $w $c
    $Output = @()
    $FieldsLength = $FieldNames.Count
    if ($FieldsLength -Lt 1) {
        Write-Host "You must specify fields to return." -ForegroundColor Red
        return
    }
    ForEach ($ListItem in $ListItems) {
        $ThisLine = ""
        $i = 0
        ForEach ($FieldName in $FieldNames) {
            $i++;
            $CleanFieldName = $FieldName.Replace(" ", "_x0020_")
            $String = $ListItem["$CleanFieldName"]
            if ($String.LookupValue) {
                $String = $String.LookupValue
            }
            if ($i -Eq $FieldsLength) {$ThisLine += "`"$String`""}
            else {$ThisLine += "`"$String`","}
        }
        $ThisLine = $ThisLine.Replace(",`n", "`n")
        $Output += $ThisLine
    }

    return $Output
}


function global:Get-AllListNames($w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns the name of every list in the SharePoint site.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $ListNames = Get-AllListNames
    #>

    $Lists = $w.Lists
    $c.Load($Lists)
    $c.ExecuteQuery()
    return $Lists
}


function global:Show-AllListNames($w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows the name of every list in the SharePoint site.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    Show-AllListNames
    #>

    $Lists = Get-AllListNames $w $c
    ForEach ($Item in $Lists) {
        Write-Host $Item.Title
    }
}


function global:Get-UserListNames($w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns the name of every user-created list in the SharePoint site.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $ListNames = Get-UserListNames
    #>

    $Filter = $global:SP_SYSTEM_LISTS
    $Lists = Get-AllListNames $w $c
    $FilteredLists = @()
    ForEach ($List in $Lists) {
        if (-Not($Filter.Contains($List.Title))) {
            $FilteredLists += $List.Title
        }
    }
    return $FilteredLists
}


function global:Show-UserListNames($w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows the name of every user-created list in the SharePoint site.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    Show-UserListNames
    #>

    $Lists = Get-UserListNames $w $c
    $Lists
}


function global:Get-AllFieldNames($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns a list of all field names in the list $ListName.
    .PARAMETER ListName
    The list whose fields should be returned.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $FieldNames = Get-AllFieldNames
    #>

    $Lists = $w.Lists
    $c.Load($Lists)
    $c.ExecuteQuery()
    $List = $Lists.getByTitle($ListName)
    $c.Load($List)
    $c.ExecuteQuery()
    $Fields = $List.Fields
    $c.Load($Fields)
    $c.ExecuteQuery()
    return $Fields
}


function global:Show-AllFieldNames($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows a list of all field names in the list $ListName.
    .PARAMETER ListName
    The list whose fields should be shown.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    Show-AllFieldNames
    #>

    $Fields = Get-AllFieldNames $ListName $w $c
    $Fields | Select Title,InternalName
}


function global:Get-UserFieldNames($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    returns a list of all user-created field names in the list $ListName.
    .PARAMETER ListName
    The list whose user-defined fields should be returned.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    $FieldNames = Get-UserFieldNames
    #>

    $Filter = $global:SP_SYSTEM_FIELDS
    $Fields = Get-AllFieldNames $ListName $w $c
    $FilteredFields = @()
    ForEach ($Field in $Fields) {
        if (-Not($Filter.Contains($Field.Title))) {
            $FilteredFields += $Field
        }
    }
    return $FilteredFields
}


function global:Show-UserFieldNames($ListName, $w=$global:SPweb, $c=$global:SPcontext) {
    <#
    .SYNOPSIS
    Shows a list of all user-created field names in the list $ListName.
    .PARAMETER ListName
    The list whose user-defined fields should be shown.
    .PARAMETER w
    The ClientContext.Web object to use to get the list. This defaults to $global:SPweb which is handled by this script.
    .PARAMETER c
    The ClientContext object to use to get the list. This defaults to $global:SPcontext which is handled by this script.
    .EXAMPLE
    Show-UserFieldNames
    #>

    $Fields = Get-UserFieldNames $ListName $w $c
    $Fields | Select Title,InternalName
}

if (!$import) {
    SPOLists-Init -silent $silent
}
