param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-hyperlinks-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    [void]$sheet.Hyperlinks.Add($sheet.Range('A1'),'https://example.com/source?x=1&y=2','','source tip','2')
    $sheet.Range('A1').Value2 = 2
    [void]$sheet.Hyperlinks.Add($sheet.Range('B1'),'','Sheet1!$J$1','internal tip','blank')
    $sheet.Range('B1').Value2 = 4
    $sheet.Range('A2').Formula = '=B2+$J$1'
    $sheet.Range('J1').Value2 = 3
    $modes = [ordered]@{all=-4104;values=-4163;formulas=-4123;formats=-4122;noBorders=7;comments=-4144;validation=6;widths=8}
    $cases = @()
    foreach ($mode in $modes.Keys) {
        foreach ($transpose in @($false,$true)) {
            foreach ($skip in @($false,$true)) {
                foreach ($operation in @('none','add')) {
                    [void]$sheet.Range('D4:E5').Clear()
                    foreach ($ref in @('D4','E4','D5','E5')) {
                        [void]$sheet.Hyperlinks.Add($sheet.Range($ref),'https://example.com/destination','','dest tip','9')
                        $sheet.Range($ref).Value2 = 9
                    }
                    [void]$sheet.Range('A1:B2').Copy()
                    $op = if ($operation -eq 'add') {2} else {-4142}
                    [void]$sheet.Range('D4').PasteSpecial($modes[$mode],$op,$skip,$transpose)
                    $cells = @()
                    foreach ($ref in @('D4','E4','D5','E5')) {
                        $cell = $sheet.Range($ref)
                        $link = $null
                        if ($cell.Hyperlinks.Count -gt 0) {
                            $h = $cell.Hyperlinks.Item(1)
                            $link = [ordered]@{target=[string]$h.Address;location=[string]$h.SubAddress;tooltip=[string]$h.ScreenTip}
                        }
                        $cells += [ordered]@{row=$cell.Row-1;col=$cell.Column-1;value=$cell.Value2;link=$link}
                    }
                    $cases += [ordered]@{mode=$mode;transpose=$transpose;skipBlanks=$skip;operation=$operation;cells=$cells}
                }
            }
        }
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 10) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) hyperlink cases"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
