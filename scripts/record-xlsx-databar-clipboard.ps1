param([string]$OutputFile = (Join-Path $env:TEMP 'xlsx-databar-clipboard.json'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Read-SheetXml([string]$Path) {
    $stream = [IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $zip = [IO.Compression.ZipArchive]::new($stream,[IO.Compression.ZipArchiveMode]::Read)
    try {
        $reader = [IO.StreamReader]::new($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
        try {return $reader.ReadToEnd()} finally {$reader.Dispose()}
    } finally {$zip.Dispose();$stream.Dispose()}
}
$excel = $null
$book = $null
$file = Join-Path $env:TEMP ('xlsx-databar-' + [guid]::NewGuid() + '.xlsx')
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $cases = @()
    foreach ($variant in @(0,1,2)) {
        foreach ($source in @('A1:A5','A2:A4')) {
            foreach ($transpose in @($false,$true)) {
                foreach ($mode in @('all','formats')) {
                    [void]$sheet.Cells.Clear()
                    for ($i = 1; $i -le 5; $i++) {$sheet.Cells.Item($i,1).Value2 = ($i-3)*10}
                    $sheet.Range('J1').Value2 = -20
                    $sheet.Range('J2').Value2 = 20
                    $bar = $sheet.Range('A1:A5').FormatConditions.AddDatabar()
                    $bar.BarFillType = if ($variant -eq 0) {1} else {0}
                    $bar.Direction = if ($variant -eq 1) {-5004} else {-5003}
                    $bar.AxisPosition = $variant
                    $bar.BarColor.Color = 65280
                    $bar.BarBorder.Type = if ($variant -eq 2) {0} else {1}
                    if ($variant -ne 2) {$bar.BarBorder.Color.Color = 255}
                    $bar.NegativeBarFormat.ColorType = if ($variant -eq 2) {1} else {0}
                    if ($variant -ne 2) {$bar.NegativeBarFormat.Color.Color = 16711680}
                    $bar.NegativeBarFormat.BorderColorType = if ($variant -eq 2) {1} else {0}
                    if ($variant -ne 2) {$bar.NegativeBarFormat.BorderColor.Color = 65535}
                    $bar.AxisColor.Color = 16711935
                    $bar.ShowValue = $variant -ne 1
                    [void]$bar.MinPoint.Modify(4,'=$J$1')
                    [void]$bar.MaxPoint.Modify(4,'=$J$2')
                    if ($cases.Count -eq 0) {$book.SaveAs($file,51)} else {$book.Save()}
                    $before = Read-SheetXml $file
                    [void]$sheet.Range($source).Copy()
                    $pasteType = if ($mode -eq 'all') {-4104} else {-4122}
                    [void]$sheet.Range('D4').PasteSpecial($pasteType,-4142,$false,$transpose)
                    $book.Save()
                    $cases += [ordered]@{variant=$variant;source=$source;transpose=$transpose;mode=$mode
                        before=$before;after=(Read-SheetXml $file)}
                }
            }
        }
    }
    [void]$sheet.Cells.Clear()
    for ($i = 1; $i -le 5; $i++) {$sheet.Cells.Item($i,1).Value2 = ($i-3)*10}
    $bar = $sheet.Range('A1:A5').FormatConditions.AddDatabar()
    [void]$bar.MinPoint.Modify(6,0)
    [void]$bar.MaxPoint.Modify(7,0)
    $bar.AxisPosition = 0
    $book.Save()
    $before = Read-SheetXml $file
    [void]$sheet.Range('A1:A5').Copy()
    [void]$sheet.Range('D4').PasteSpecial(-4104)
    $book.Save()
    $cases += [ordered]@{variant=3;source='A1:A5';transpose=$false;mode='all'
        before=$before;after=(Read-SheetXml $file)}
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputFile),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 5) +
        [Environment]::NewLine,[Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) native data-bar copies"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
    if (Test-Path -LiteralPath $file) {Remove-Item -LiteralPath $file}
}
