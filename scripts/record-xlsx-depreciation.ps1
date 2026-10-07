# Native DB/DDB differential fixture recorder; uses its own hidden workbook.
param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-depreciation-native.json'))
$ErrorActionPreference = 'Stop'
$formulas = @()
foreach ($factor in @(0.5,1,1.5,2,10,20)) {
    foreach ($period in @(0.1,0.5,1,1.5,2,2.5,8.5,9,9.5,10)) {
        $formulas += "=DDB(2400,300,10,$period,$factor)"
    }
}
foreach ($life in @(0.5,1,1.5,10,1000000000000)) {
    $formulas += "=DDB(2400,300,$life,$life,2)"
}
foreach ($cost in @(-1,0,2400)) {
    foreach ($salvage in @(-1,0,2390,2400,2500)) {
        $formulas += "=DDB($cost,$salvage,10,1)"
        $formulas += "=DDB($cost,$salvage,10,9.5)"
    }
}
$formulas += @(
    '=DDB(2400,300,10,0)', '=DDB(2400,300,10,-0.5)',
    '=DDB(2400,300,10,11)', '=DDB(2400,300,0,1)',
    '=DDB(2400,300,-1,1)', '=DDB(2400,300,10,1,0)',
    '=DDB(2400,300,10,1,-1)', '=DDB(2400,300,10,1,)',
    '=DDB(2400,300,10,1,1/0)', '=DDB(2400,300,10,1,"bad")',
    '=DDB(2400,300,10,1,"2")', '=DDB(2400,300,10,1,TRUE)',
    '=DDB(2400,300,10,1,FALSE)'
)
foreach ($month in @(1,7,7.9,11.9,12,12.9)) {
    foreach ($period in @(0.5,1,1.5,2,2.5,10,10.5,10.9,11,11.9)) {
        $formulas += "=DB(2400,300,10,$period,$month)"
    }
}
foreach ($life in @(0.5,1,1.5,10.9,1000000000000)) {
    $formulas += "=DB(2400,300,$life,$life,7)"
}
$formulas += @(
    '=DB(0,300,10,1)', '=DB(2400,2500,10,1)', '=DB(2400,300,10,0)',
    '=DB(2400,300,10,-0.5)', '=DB(2400,300,0,1)', '=DB(2400,300,10,1,0.9)',
    '=DB(2400,300,10,1,13)', '=DB(2400,300,10,1,)', '=DB(2400,300,10,1,"bad")',
    '=DB(2400,300,10,1,1/0)', '=DB(-1,300,10,1)', '=DB(2400,-1,10,1)'
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
    Write-Output ('Excel ' + $excel.Version + ': ' + $results.Count + ' DB/DDB cases recorded')
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
