# Records formula results from a new, hidden native Excel instance.
# Intentional test fixture generation; it never attaches to a user's workbook.
param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-amorlinc-native.json'))
$ErrorActionPreference = 'Stop'
$formulas = @()
foreach ($basis in @(0,1,3,4)) {
    foreach ($period in @(0,1,5,6,7,100,1.9)) {
        $formulas += "=AMORLINC(2400,39679,39813,300,$period,0.15,$basis)"
    }
    foreach ($dates in @(@(39679,39679),@(39679,40045),@(43890,43921),@(59,61),@(0,1),@(43889,43890),@(43890,43891),@(43466,43890))) {
        $formulas += "=AMORLINC(2400,$($dates[0]),$($dates[1]),300,0,0.15,$basis)"
    }
}
foreach ($cost in @(0,-1,2400)) {
    foreach ($salvage in @(-1,0,2390,2400,2500)) {
        foreach ($period in @(0,1)) {
            $formulas += "=AMORLINC($cost,39679,39813,$salvage,$period,0.15)"
        }
    }
}
foreach ($rate in @(-0.1,0,0.15,2,10)) {
    foreach ($period in @(0,1,6,-0.5)) {
        $formulas += "=AMORLINC(2400,39679,39813,300,$period,$rate)"
    }
}
foreach ($basis in @(-1,-0.5,1.9,2,4.9,5)) {
    $formulas += "=AMORLINC(2400,39679,39813,300,0,0.15,$basis)"
}
$formulas += @(
    '=AMORLINC(2400,DATE(2019,12,31),DATE(2020,3,31),300,0,0.15,1)',
    '=AMORLINC(2400,DATE(2020,12,31),DATE(2021,3,31),300,0,0.15,1)',
    '=AMORLINC(2400,DATE(2019,1,1),DATE(2021,1,1),300,0,0.15,1)',
    '=AMORLINC(2400,39814,39813,300,0,0.15)',
    '=AMORLINC(2400,-0.1,1,300,0,0.15)',
    '=AMORLINC(2400,39679.9,39813.1,300,0,0.15)',
    '=AMORLINC(2400,39679,39813,300,1,0.15,)',
    '=AMORLINC(2400,39679,39813,300,,0.15)',
    '=AMORLINC(2400,39679,39813,,0,0.15)',
    '=AMORLINC(2400,39679,39813,300,A1,0.15)',
    '=AMORLINC(2400,39679,39813,A1,0,0.15)',
    '=AMORLINC(2400,39679,39813,300,0,"0.15")',
    '=AMORLINC(2400,39679,39813,300,0,"bad")',
    '=AMORLINC(2400,39679,39813,300,0,1/0)'
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    $sheet.Columns('Z').ColumnWidth = 120
    $cell = $sheet.Range('Z1')
    $results = @()
    foreach ($formula in ($formulas | Select-Object -Unique)) {
        $cell.Formula2 = $formula
        $excel.Calculate()
        $error = $sheet.Evaluate('ISERROR(Z1)')
        $value = if ($error) { [string]$cell.Text } else { $cell.Value2 }
        $results += @{ formula = $formula; value = $value; type = $(if ($error) { 'error' } else { 'number' }) }
    }
    @{ version = $excel.Version; cases = $results } | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath $OutputFile -Encoding utf8
    Write-Output ('Excel ' + $excel.Version + ': ' + $results.Count + ' cases recorded to ' + $OutputFile)
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
