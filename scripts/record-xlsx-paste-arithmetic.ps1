param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-arithmetic-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $pairs = @(
        @('2','10'), @('','10'), @('2',''), @('',''), @("'2",'10'),
        @('text','10'), @('2','text'), @('TRUE','10'), @('2','FALSE'),
        @('=1+1','10'), @('2','=5+5'), @('=1+1','=5+5'),
        @('=""','10'), @('2','=""'), @('=1/0','10'), @('2','=NA()'),
        @('0','10'), @('2',"'10"), @('=TRUE()','10'), @('="2"','10'),
        @('="text"','10'), @('-2','=5+5'), @('=1+1','-10'), @("'1,000",'=5+5'),
        @('=1+1',"'1,000"), @('','=""'), @('0.1','0.2'), @('1.2','1.2'),
        @('#N/A','10'), @('2','#N/A'), @('=1+1','TRUE'), @('="text"','text')
    )
    $cases = @()
    foreach ($mode in @(@{name='all';code=-4104},@{name='values';code=-4163},
        @{name='formulas';code=-4123},@{name='formats';code=-4122})) {
        foreach ($op in @(@{name='add';code=2},@{name='subtract';code=3},
            @{name='multiply';code=4},@{name='divide';code=5})) {
            foreach ($pair in $pairs) {
                [void]$sheet.Range('A1:C1').Clear()
                if ($pair[0] -ne '') { $sheet.Range('A1').Formula = $pair[0] }
                if ($pair[1] -ne '') { $sheet.Range('C1').Formula = $pair[1] }
                $sheet.Range('A1').Font.Bold = $true
                $sheet.Range('C1').Font.Italic = $true
                [void]$sheet.Range('A1').Copy()
                [void]$sheet.Range('C1').PasteSpecial($mode.code,$op.code,$false,$false)
                $excel.Calculate()
                $cell = $sheet.Range('C1')
                $value = $cell.Value2
                if ($cell.Text -match '^#') { $value = @{error=$cell.Text} }
                $cases += [ordered]@{
                    mode=$mode.name; operation=$op.name; source=$pair[0]; destination=$pair[1]
                    value=$value; formula=$(if ($cell.HasFormula) {$cell.Formula.Substring(1)} else {$null})
                    bold=[bool]$cell.Font.Bold; italic=[bool]$cell.Font.Italic
                }
            }
        }
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 8) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) paste arithmetic cases"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
}
