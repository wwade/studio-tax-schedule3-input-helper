param(
    [Parameter(Position = 0)]
    [string]$CsvPath = ".\schedule3.csv",
    [switch]$Preview,
    [int]$MaxRows = 0,
    [int]$CountdownSeconds = 5,
    [int]$ExtraTabsAfterRow = 1,
    [int]$DelayMs = 80
)

$ErrorActionPreference = 'Stop'

$columns = @(
    'Number',
    'Name of fund/corp.',
    'Year of acquisition',
    'Proceeds of disposition',
    'Adjusted cost base',
    'Outlays and expenses'
)

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV not found: $CsvPath"
}

$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows -or $rows.Count -eq 0) {
    throw "CSV has no data rows: $CsvPath"
}
if ($MaxRows -gt 0) {
    $rows = @($rows | Select-Object -First $MaxRows)
}

$headers = @($rows[0].PSObject.Properties.Name)
$missing = @($columns | Where-Object { $_ -notin $headers })
if ($missing.Count -gt 0) {
    throw "CSV is missing required column(s): $($missing -join ', ')"
}

function Format-StudioTaxValue {
    param(
        [object]$Value,
        [string]$Column
    )

    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ($Column -in @('Number', 'Year of acquisition')) {
        return $text.Trim()
    }
    if ($Column -in @('Proceeds of disposition', 'Adjusted cost base', 'Outlays and expenses')) {
        $number = 0.0
        if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
            return $number.ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $text.Trim()
}

function Send-GridText {
    param([string]$Text)

    # Escape characters that SendKeys treats as commands.
    $escaped = $Text
    $escaped = $escaped -replace '\+', '{+}'
    $escaped = $escaped -replace '\^', '{^}'
    $escaped = $escaped -replace '%', '{%}'
    $escaped = $escaped -replace '~', '{~}'
    $escaped = $escaped -replace '\(', '{(}'
    $escaped = $escaped -replace '\)', '{)}'
    $escaped = $escaped -replace '\[', '{[}'
    $escaped = $escaped -replace '\]', '{]}'
    $escaped = $escaped -replace '\{', '{{}'
    $escaped = $escaped -replace '\}', '{}}'

    [System.Windows.Forms.SendKeys]::SendWait($escaped)
}

function Write-Summary {
    Write-Host "Prepared $($rows.Count) Schedule 3 row(s)."
    Write-Host ("Expected Proceeds total: {0:0.00}" -f $summary['Proceeds of disposition'])
    Write-Host ("Expected Gain/Loss total: {0:0.00}" -f $summary['Gain or Loss'])
}

$expected = $rows | Measure-Object -Property 'Proceeds of disposition', 'Adjusted cost base', 'Outlays and expenses', 'Gain or Loss' -Sum
$summary = @{}
foreach ($item in $expected) {
    $summary[$item.Property] = $item.Sum
}

Write-Summary
Write-Host ''
Write-Host 'First two rows to enter:'
foreach ($row in ($rows | Select-Object -First 2)) {
    $previewValues = foreach ($column in $columns) {
        Format-StudioTaxValue -Value $row.$column -Column $column
    }
    Write-Host ($previewValues -join ' | ')
}

if ($Preview) {
    exit 0
}

Add-Type -AssemblyName System.Windows.Forms

Write-Host ''
Write-Host "Click the first yellow Number cell in StudioTax now. Cell-by-cell entry starts in $CountdownSeconds second(s)..."
for ($i = $CountdownSeconds; $i -gt 0; $i--) {
    Write-Host "$i..."
    Start-Sleep -Seconds 1
}

$rowNumber = 0
foreach ($row in $rows) {
    $rowNumber++
    Write-Host "Entering row $rowNumber of $($rows.Count): $($row.'Name of fund/corp.')"

    foreach ($column in $columns) {
        $value = Format-StudioTaxValue -Value $row.$column -Column $column
        Send-GridText -Text $value
        Start-Sleep -Milliseconds $DelayMs
        [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
        Start-Sleep -Milliseconds $DelayMs
    }

    for ($i = 0; $i -lt $ExtraTabsAfterRow; $i++) {
        [System.Windows.Forms.SendKeys]::SendWait('{TAB}')
        Start-Sleep -Milliseconds $DelayMs
    }
}

Write-Host ''
Write-Host 'Cell-by-cell entry finished. Review totals in StudioTax before clicking OK.'
Write-Summary
