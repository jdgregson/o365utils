Param (
    [string]
    $UserIds,

    [System.ValueType]
    $StartDate = (Date).AddDays(-3),

    [System.ValueType]
    $EndDate = (Date).AddDays(1)
)

$deletions = (Search-UnifiedAuditLog `
    -StartDate $StartDate `
    -EndDate $EndDate `
    -ResultSize 5000 `
    -UserIds $UserIds `
    -Operations 'SoftDelete','MoveToDeletedItems'
)

$output = @()
$global:truncatedRecords = @()
$deletions | ForEach-Object {
    $deletion = $_
    try {
        $auditData = ConvertFrom-JSON $deletion.AuditData
        $auditData.AffectedItems | ForEach-Object {
            $output += New-Object PSObject -Property @{
                Operation = $deletion.Operations
                Subject = $_.Subject
                Path = $_.ParentFolder.Path
            }
        }
    } catch {
        $global:truncatedRecords += $deletion
        Write-Warning "Record truncated, thanks Microsoft! Saved in `$global:truncatedRecords."
    }
}
$output
