# Records clipboard behavior in a fresh hidden Excel instance, without user workbooks.
param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-options-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)
    $sheet.Range('A1').Font.Bold = $true
    $sheet.Range('A1').Interior.Color = 255
    $sheet.Range('A2').Formula = '=""'
    $sheet.Range('A3').Value2 = 0
    $sheet.Range('A4').Value2 = $false
    $cases = @()
    foreach ($mode in @(@{name='all';code=-4104},@{name='values';code=-4163},
        @{name='formulas';code=-4123},@{name='formats';code=-4122})) {
        foreach ($transpose in @($false, $true)) {
            foreach ($skipBlanks in @($false, $true)) {
                [void]$sheet.Range('C1:G5').Clear()
                $dest = $sheet.Range($(if ($transpose) {'C1:G1'} else {'C1:C5'}))
                $dest.Value2 = 9
                $dest.Font.Bold = $false
                $dest.Font.Italic = $true
                $dest.Interior.Color = 65280
                [void]$sheet.Range('A1:A5').Copy()
                [void]$sheet.Range('C1').PasteSpecial($mode.code, -4142, $skipBlanks, $transpose)
                $excel.Calculate()
                $cells = @()
                for ($i = 0; $i -lt 5; $i++) {
                    $cell = $sheet.Cells.Item($(if ($transpose) {1} else {$i + 1}),
                        $(if ($transpose) {$i + 3} else {3}))
                    $fill = $null
                    if ($cell.Interior.Pattern -eq 1) {
                        $color = [int]$cell.Interior.Color
                        $fill = '{0:X2}{1:X2}{2:X2}' -f ($color -band 255),
                            (($color -shr 8) -band 255), (($color -shr 16) -band 255)
                    }
                    $cells += [ordered]@{
                        value = $cell.Value2
                        formula = $(if ($cell.HasFormula) {$cell.Formula.Substring(1)} else {$null})
                        bold = [bool]$cell.Font.Bold
                        italic = [bool]$cell.Font.Italic
                        fill = $fill
                    }
                }
                $cases += [ordered]@{mode=$mode.name; transpose=$transpose; skipBlanks=$skipBlanks; cells=$cells}
            }
        }
    }
    $result = [ordered]@{version=$excel.Version; build=$excel.Build; cases=$cases}
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ($result | ConvertTo-Json -Depth 8) + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) paste cases with Excel $($excel.Version) to $OutputFile"
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
