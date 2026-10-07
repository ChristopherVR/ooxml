param(
    [string]$OutputFolder = (Join-Path $env:TEMP 'xlsx-databar-geometry'),
    [ValidateRange(0,100)][int]$PercentMin = 0,
    [ValidateRange(0,100)][int]$PercentMax = 100,
    [switch]$Context,
    [ValidateSet(-5002,-5003,-5004)][int]$ReadingOrder = -5002,
    [ValidateSet('mixed','positive','negative','zero','equalPositive','equalNegative')][string[]]$Kinds = @('mixed','positive','negative','zero','equalPositive','equalNegative')
)
$ErrorActionPreference = 'Stop'
if ($PercentMax -lt $PercentMin) {throw 'PercentMax must be at least PercentMin'}
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Directory]::CreateDirectory($OutputFolder) | Out-Null
function Read-SheetXml([string]$Path, [string]$Part = 'xl/worksheets/sheet1.xml') {
    $stream = [IO.File]::Open($Path, 'Open', 'Read', 'ReadWrite')
    $zip = [IO.Compression.ZipArchive]::new($stream, 'Read')
    try {
        $reader = [IO.StreamReader]::new($zip.GetEntry($Part).Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose(); $stream.Dispose() }
}
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add()
    $sheet = $book.Sheets.Item(1)
    $sheet.Columns.Item(1).ColumnWidth = 12
    $sheet.Columns.Item(2).ColumnWidth = 30
    $sheet.PageSetup.PrintArea = '$A$1:$B$8'
    $sheet.PageSetup.Zoom = 100
    $sheet.PageSetup.LeftMargin = 36
    $sheet.PageSetup.TopMargin = 36
    $cases = @()
    foreach ($kind in $Kinds) {
        $values = switch ($kind) {
            'mixed' { @(-20,-10,0,10,40) }
            'positive' { @(10,20,30,40,50) }
            'negative' { @(-50,-40,-30,-20,-10) }
            'zero' { @(0,0,0,0,0) }
            'equalPositive' { @(10,10,10,10,10) }
            'equalNegative' { @(-10,-10,-10,-10,-10) }
        }
        foreach ($axis in @(0,1,2)) {
            foreach ($auto in @($true,$false)) {
                foreach ($rtl in @($false,$true)) {
                    [void]$sheet.Cells.Clear()
                    $id = "$kind-$axis-$auto-$rtl"
                    if ($PercentMin -ne 0 -or $PercentMax -ne 100) {$id += "-lengths-$PercentMin-$PercentMax"}
                    if ($Context) {$id += "-context-$ReadingOrder"}
                    $sheet.DisplayRightToLeft = $Context -and $rtl
                    $sheet.Range('A1').Value2 = $id
                    # A yellow reference cell lets the extractor normalize printed coordinates.
                    $sheet.Range('B8').Interior.Color = 65535
                    $sheet.Range('B8').Value2 = 1.0
                    $referenceBar = $sheet.Range('B8').FormatConditions.AddDatabar()
                    $referenceBar.ShowValue = $false
                    $referenceBar.BarFillType = 0
                    $referenceBar.BarColor.Color = 65280
                    $referenceBar.AxisPosition = 2
                    [void]$referenceBar.MinPoint.Modify(0,0)
                    [void]$referenceBar.MaxPoint.Modify(0,1)
                    for ($i = 0; $i -lt $values.Count; $i++) {
                        $sheet.Cells.Item($i+2,1).Value2 = [double]$values[$i]
                        $sheet.Cells.Item($i+2,2).Value2 = [double]$values[$i]
                    }
                    $sheet.Range('B2:B6').ReadingOrder = $ReadingOrder
                    $bar = $sheet.Range('B2:B6').FormatConditions.AddDatabar()
                    $bar.ShowValue = $false
                    $bar.BarFillType = 0
                    $bar.BarColor.Color = 65280
                    $bar.AxisPosition = $axis
                    $bar.PercentMin = $PercentMin
                    $bar.PercentMax = $PercentMax
                    $bar.AxisColor.Color = 16711935
                    $bar.Direction = if ($Context) {-5002} elseif ($rtl) {-5004} else {-5003}
                    $bar.NegativeBarFormat.ColorType = 0
                    $bar.NegativeBarFormat.Color.Color = 16711680
                    if ($auto) {
                        [void]$bar.MinPoint.Modify(6,0)
                        [void]$bar.MaxPoint.Modify(7,0)
                    } else {
                        [void]$bar.MinPoint.Modify(0,($values | Measure-Object -Minimum).Minimum)
                        [void]$bar.MaxPoint.Modify(0,($values | Measure-Object -Maximum).Maximum)
                    }
                    $file = Join-Path $OutputFolder 'native.xlsx'
                    if ($cases.Count -eq 0) {$book.SaveAs($file,51)} else {$book.Save()}
                    $sheet.ExportAsFixedFormat(0,(Join-Path $OutputFolder "$id.pdf"))
                    $record = [ordered]@{id=$id;values=$values;axis=$axis;auto=$auto;rtl=$rtl;percentMin=$PercentMin;percentMax=$PercentMax;context=[bool]$Context;readingOrder=$ReadingOrder;xml=(Read-SheetXml $file)}
                    if ($Context) {$record.styles = Read-SheetXml $file 'xl/styles.xml'}
                    $cases += $record
                }
            }
        }
    }
    [IO.File]::WriteAllText((Join-Path $OutputFolder 'raw.json'),
        ([ordered]@{version=$excel.Version;build=$excel.Build;cases=$cases} | ConvertTo-Json -Depth 6),
        [Text.UTF8Encoding]::new($false))
    Write-Output "Recorded $($cases.Count) native PDF and worksheet geometry probes in $OutputFolder"
} finally {
    if ($null -ne $book) {$book.Close($false)}
    if ($null -ne $excel) {$excel.Quit()}
}
