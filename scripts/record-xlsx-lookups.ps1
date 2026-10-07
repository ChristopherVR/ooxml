# Native lookup differential recorder: always owns a hidden temporary workbook.
param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-lookups-native.json'))
$ErrorActionPreference = 'Stop'
$lookupValues = @('A7', '""', '0', '2', '4', '"b"', 'TRUE', 'FALSE', '"a*"')
$scenarios = @(
    @{ name = 'ascending mixed values and blank tail'; cells = @{ A1 = 1; A2 = 3; A3 = 'a'; A4 = $false } },
    @{ name = 'descending mixed values and blank head'; cells = @{ A3 = $false; A4 = 'a'; A5 = 3; A6 = 1 } }
)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $null
try {
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Columns('Z').ColumnWidth = 120
    $cell = $sheet.Range('Z1')
    $results = @()
    foreach ($scenario in $scenarios) {
        $sheet.Range('A1:A7').ClearContents()
        foreach ($address in $scenario.cells.Keys) {
            $inputValue = $scenario.cells[$address]
            $inputCell = $sheet.Range([string]$address)
            $inputCell.Formula2 = if ($inputValue -is [string]) {
                '="' + $inputValue.Replace('"', '""') + '"'
            } elseif ($inputValue -is [bool]) {
                if ($inputValue) { '=TRUE()' } else { '=FALSE()' }
            } else {
                '=' + $inputValue.ToString([Globalization.CultureInfo]::InvariantCulture)
            }
        }
        foreach ($lookup in $lookupValues) {
            foreach ($matchMode in @(0,-1,1,2)) {
                # Binary search requires its declared sort order.
                $binaryMode = if ($scenario.name.StartsWith('ascending')) { 2 } else { -2 }
                foreach ($searchMode in @(1,-1,$binaryMode)) {
                    $formula = "XMATCH($lookup,A1:A6,$matchMode,$searchMode)"
                    $cell.Formula2 = "=$formula"
                    $excel.Calculate()
                    $isError = $sheet.Evaluate('ISERROR(Z1)')
                    $value = if ($isError) { [string]$cell.Text } else { $cell.Value2 }
                    $results += @{
                        scenario = $scenario.name; cells = $scenario.cells; formula = $formula
                        value = $value; type = $(if ($isError) { 'error' } else { 'number' })
                    }
                }
            }
        }
    }
    @{ version = $excel.Version; build = $excel.Build; cases = $results } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputFile -Encoding utf8
    Write-Output ('Excel ' + $excel.Version + ': ' + $results.Count + ' lookup cases recorded')
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}
