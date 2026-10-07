param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-borders-native.json'))
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
    $sheet.Range('A2').Formula = '=B2+$J$1'
    $sheet.Range('B1').Value2 = 'text'
    $sheet.Range('J1').Value2 = 3
    $source = $sheet.Range('A1:B2')
    $source.Font.Bold = $true
    $source.NumberFormat = '0.00'
    $source.Interior.Color = 255
    $edges = @(@{name='left';code=7},@{name='right';code=10},@{name='top';code=8},
        @{name='bottom';code=9},@{name='diagonalDown';code=5},@{name='diagonalUp';code=6})
    foreach ($edge in $edges) {
        $source.Borders.Item($edge.code).LineStyle = 1
        $source.Borders.Item($edge.code).Weight = 2
        $source.Borders.Item($edge.code).Color = 255
    }
    $cases = @()
    foreach ($op in @(@{name='none';code=-4142},@{name='add';code=2},@{name='subtract';code=3},
        @{name='multiply';code=4},@{name='divide';code=5})) {
        foreach ($transpose in @($false,$true)) {
            foreach ($skip in @($false,$true)) {
                $dest = $sheet.Range('D4:E5')
                [void]$dest.Clear()
                $dest.Value2 = 9
                $dest.Font.Italic = $true
                $dest.Interior.Color = 65280
                foreach ($cell in $dest.Cells) {
                    foreach ($edge in $edges) {
                        $cell.Borders.Item($edge.code).LineStyle = -4119
                        $cell.Borders.Item($edge.code).Color = 16711680
                    }
                }
                [void]$source.Copy()
                [void]$sheet.Range('D4').PasteSpecial(7,$op.code,$skip,$transpose)
                $excel.Calculate()
                $cells = @()
                foreach ($cell in $dest.Cells) {
                    $value = $cell.Value2
                    if ($cell.Text -match '^#') { $value = @{error=$cell.Text} }
                    $borders = [ordered]@{}
                    foreach ($edge in $edges) {
                        $b = $cell.Borders.Item($edge.code)
                        $color = [int]$b.Color
                        $borders[$edge.name] = [ordered]@{
                            style=$b.LineStyle
                            color=('{0:X2}{1:X2}{2:X2}' -f ($color -band 255),
                                (($color -shr 8) -band 255),(($color -shr 16) -band 255))
                        }
                    }
                    $cells += [ordered]@{row=$cell.Row-4;col=$cell.Column-4;value=$value
                        formula=$(if ($cell.HasFormula) {$cell.Formula.Substring(1)} else {$null})
                        bold=[bool]$cell.Font.Bold;italic=[bool]$cell.Font.Italic
                        numFmt=$cell.NumberFormat;fill=$cell.Interior.Color;borders=$borders}
                }
                $cases += [ordered]@{mode='noBorders';operation=$op.name;transpose=$transpose;skipBlanks=$skip;cells=$cells}
            }
        }
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 10) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) All Except Borders cases"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
