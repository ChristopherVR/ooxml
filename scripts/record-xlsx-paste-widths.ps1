param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-paste-widths-native.json'))
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $sheet.Range('A:A').ColumnWidth = 20
    $sheet.Range('B:B').ColumnWidth = 30
    $sheet.Columns.Item(2).Hidden = $true
    $sheet.Range('A1').Value2 = 2
    $sheet.Range('B1').Value2 = 3
    $cases = @()
    foreach ($source in @('A1:B1','A1:A3','C1:D2','A1:C2','A:B')) {
        foreach ($dest in @('F4','F4:I5','F4:I6','F:I')) {
            foreach ($transpose in @($false,$true)) {
                foreach ($skip in @($false,$true)) {
                    foreach ($i in 6..9) {
                        $sheet.Columns.Item($i).ColumnWidth = 12
                        $sheet.Columns.Item($i).Hidden = ($i % 2 -eq 0)
                    }
                    $sheet.Range('F4:I6').Value2 = 9
                    [void]$sheet.Range($source).Copy()
                    $errorText = $null
                    try { [void]$sheet.Range($dest).PasteSpecial(8,2,$skip,$transpose) }
                    catch { $errorText = $_.Exception.Message }
                    $columns = @()
                    foreach ($i in 6..9) {
                        $column = $sheet.Columns.Item($i)
                        $hidden = [bool]$column.Hidden
                        $width = $column.ColumnWidth
                        if ($hidden) {$column.Hidden = $false}
                        $storedWidth = $column.ColumnWidth
                        $column.Hidden = $hidden
                        $columns += [ordered]@{width=$width;hidden=$hidden;storedWidth=$storedWidth}
                    }
                    $cases += [ordered]@{source=$source;dest=$dest;transpose=$transpose;skipBlanks=$skip
                        operation='add';error=$errorText;columns=$columns;value=$sheet.Range('F4').Value2}
                }
            }
        }
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;defaultWidth=$sheet.StandardWidth;cases=$cases} | ConvertTo-Json -Depth 8) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) Column Widths cases"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
