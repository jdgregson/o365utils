# File: Block-AttachmentExtensions.ps1
# Purpose: this script will connect to Exchange Online and block a list of
# abused file extensions.
#
# Copyright (c) 2019, jdgregson
# Author: Jonathan Gregson <jonathan@jdgregson.com>
#                          <jdgregson@gmail.com>

. O365-Auth.ps1

$DefaultRuleName = "Blocked Attachments"
$DefaultBlockedExtensions = "ace",
    "ade",
    "adp",
    "ani",
    "app",
    "appcontent-ms",
    "asp",
    "bas",
    "bat",
    "bin",
    "cer",
    "chm",
    "cla",
    "class",
    "cmd",
    "cnt",
    "com",
    "cpl",
    "crt",
    "csh",
    "der",
    "diagcab",
    "dll",
    "docm",
    "dos",
    "email",
    "exe",
    "fxp",
    "gadget",
    "grp",
    "hlp",
    "hpj",
    "hta",
    "html",
    "img",
    "inf",
    "ins",
    "iso",
    "isp",
    "its",
    "jar",
    "jnlp",
    "js",
    "jse",
    "ksh",
    "lnk",
    "mad",
    "maf",
    "mag",
    "mam",
    "maq",
    "mar",
    "mas",
    "mat",
    "mau",
    "mav",
    "maw",
    "mcf",
    "mda",
    "mdb",
    "mde",
    "mdt",
    "mdw",
    "mdz",
    "ms",
    "msc",
    "msh",
    "msh1",
    "msh1xml",
    "msh2",
    "msh2xml",
    "mshxml",
    "msi",
    "msp",
    "mst",
    "msu",
    "obj",
    "ocx",
    "ops",
    "os2",
    "osd",
    "pcd",
    "pif",
    "pl",
    "plg",
    "prf",
    "prg",
    "printerexport",
    "ps1",
    "ps1xml",
    "ps2",
    "ps2xml",
    "psc1",
    "psc2",
    "psd1",
    "psdm1",
    "pst",
    "py",
    "pyc",
    "pyo",
    "pyw",
    "pyz",
    "pyzw",
    "rar",
    "rdp",
    "reg",
    "rtf",
    "scf",
    "sct",
    "settingcontent-ms",
    "shb",
    "shs",
    "theme",
    "tmp",
    "url",
    "vb",
    "vbe",
    "vbp",
    "vbp",
    "vbs",
    "vhd",
    "vhdx",
    "vsmacros",
    "vsw",
    "vxd",
    "w16",
    "webpnp",
    "website",
    "ws",
    "wsc",
    "wsf",
    "wsh",
    "xbap",
    "xll",
    "xlsb",
    "xlsm",
    "xml",
    "xnk"

O365-Auth -Exchange

$existingRule = Get-TransportRule $DefaultRuleName -ErrorAction SilentlyContinue
if ($existingRule -eq $null) {
    Write-Host "A rule named `"$DefaultRuleName`" was not found. Would you like to create it?"
    $confirm = Read-Host "[Y] Yes  [N] No  (default is `"N`")"
    if ($confirm -eq "y") {
        $newRule = New-TransportRule -Name $DefaultRuleName `
            -FromScope "NotInOrganization" `
            -AttachmentExtensionMatchesWords $DefaultBlockedExtensions `
            -SetAuditSeverity "High" `
            -Quarantine $true `
            -StopRuleProcessing $true
        Write-Host "The rule `"$DefaultRuleName`" has been created."
    } else {
        Write-Host "No rules were created or modified."
        exit
    }
} else {
    Set-TransportRule $DefaultRuleName -AttachmentExtensionMatchesWords $DefaultBlockedExtensions
    Write-Host "Extension block list has been applied to the rule `"$DefaultRuleName`"."
}
