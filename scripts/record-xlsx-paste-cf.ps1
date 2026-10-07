param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-cf-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $sheet.Range('A1').Value2 = 2
    $sheet.Range('B1').Value2 = 4
    $sheet.Range('A2').Formula = '=B2+$J$1'
    $sheet.Range('J1').Value2 = 3
    $modes = [ordered]@{all=-4104;values=-4163;formulas=-4123;formats=-4122;noBorders=7;comments=-4144;validation=6;widths=8;mergeFormats=14}
    $cases = @()
    foreach ($sourceRange in @('A1:B2','A2:B3','Z1:AA2')) {
    foreach ($mode in $modes.Keys) {
        foreach ($transpose in @($false,$true)) {
            foreach ($skip in @($false,$true)) {
                foreach ($operation in @('none','add')) {
                    [void]$sheet.Cells.FormatConditions.Delete()
                    [void]$sheet.Range('D3:F6').Clear()
                    $sheet.Range('D4:E5').Value2 = 9
                    $source = $sheet.Range('A1:B2').FormatConditions.Add(2,1,'=A1>0')
                    $source.Interior.Color = 255
                    $source.Font.Bold = $true
                    $source.StopIfTrue = $false
                    $second = $sheet.Range('A1:B1').FormatConditions.Add(1,5,'=1')
                    $second.Interior.Color = 65535
                    $second.StopIfTrue = $true
                    $outside = $sheet.Range('D3:F6').FormatConditions.Add(2,1,'=D3<0')
                    $outside.Interior.Color = 16711680
                    $outside.StopIfTrue = $false
                    $dest = $sheet.Range('D4:E5').FormatConditions.Add(2,1,'=D4>5')
                    $dest.Interior.Color = 65280
                    $dest.StopIfTrue = $false
                    $source.Priority = 1
                    $second.Priority = 2
                    $outside.Priority = 3
                    $dest.Priority = 4
                    [void]$sheet.Range($sourceRange).Copy()
                    $op = if ($operation -eq 'add') {2} else {-4142}
                    [void]$sheet.Range('D4').PasteSpecial($modes[$mode],$op,$skip,$transpose)
                    $rules = @()
                    foreach ($cf in $sheet.Cells.FormatConditions) {
                        $operator = $null
                        if ($cf.Type -eq 1) {$operator=$cf.Operator}
                        $rules += [ordered]@{type=$cf.Type;formula1=[string]$cf.Formula1;operator=$operator;priority=$cf.Priority
                            ranges=$cf.AppliesTo.Address();fill=$cf.Interior.Color;bold=($cf.Font.Bold -eq $true);stopIfTrue=$cf.StopIfTrue}
                    }
                    $cells = @()
                    foreach ($ref in @('D4','E4','D5','E5','D3','F4','F6')) {
                        $cell = $sheet.Range($ref)
                        $cells += [ordered]@{row=$cell.Row-1;col=$cell.Column-1;value=$cell.Value2;fill=$cell.DisplayFormat.Interior.Color
                            priorities=@($cell.FormatConditions | ForEach-Object {$_.Priority})}
                    }
                    $cases += [ordered]@{source=$sourceRange;mode=$mode;transpose=$transpose;skipBlanks=$skip;operation=$operation;rules=$rules;cells=$cells}
                }
            }
        }
    }
    }
    [void]$sheet.Cells.Clear()
    $sheet.Range('A1').Value2 = 0
    $sheet.Range('B1').Value2 = 10
    $scale = $sheet.Range('A1:B1').FormatConditions.AddColorScale(2)
    $scale.ColorScaleCriteria.Item(1).FormatColor.Color = 255
    $scale.ColorScaleCriteria.Item(2).FormatColor.Color = 65280
    $scaleCells = @()
    $refs = @('D4','E4','F4','G4','D5','E5','F5','G5')
    for ($i = 0; $i -lt $refs.Count; $i++) {$sheet.Range($refs[$i]).Value2 = $i * 10}
    [void]$sheet.Range('A1:B1').Copy()
    [void]$sheet.Range('D4:G5').PasteSpecial(-4122)
    foreach ($ref in $refs) {
        $cell = $sheet.Range($ref)
        $scaleCells += [ordered]@{ref=$ref;value=$cell.Value2;fill=$cell.DisplayFormat.Interior.Color}
    }
    $colorScale = [ordered]@{cells=$scaleCells;ranges=@($sheet.Cells.FormatConditions | ForEach-Object {$_.AppliesTo.Address()})}
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases;colorScale=$colorScale} | ConvertTo-Json -Depth 10) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) conditional format cases"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
