param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-gradient-presets'))
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2=[string]'Month'; $sheet.Cells.Item(1,2).Value2=[string]'Sales'
    $sheet.Cells.Item(2,1).Value2=[string]'Jan'; $sheet.Cells.Item(2,2).Value2=[double]10
    $sheet.Cells.Item(3,1).Value2=[string]'Feb'; $sheet.Cells.Item(3,2).Value2=[double]20
    $chart=$sheet.ChartObjects().Add(200,20,480,300).Chart; $chart.ChartType=51
    $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $cases=@()
    foreach($id in 1..24) {
        $chart.SeriesCollection(1).Format.Fill.PresetGradient(1,1,$id)
        $path=Join-Path $OutputFolder "preset-$id.xlsx"; $book.SaveCopyAs($path)
        $zip=[IO.Compression.ZipFile]::OpenRead($path)
        try {
            $reader=[IO.StreamReader]::new($zip.GetEntry('xl/charts/chart1.xml').Open())
            try {$chartXml=$reader.ReadToEnd()} finally {$reader.Dispose()}
        } finally {$zip.Dispose()}
        $probe=$excel.Workbooks.Open($path,0,$true)
        try {
            $fill=$probe.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Format.Fill
            $stops=@(); foreach($n in 1..$fill.GradientStops.Count) {
                $stop=$fill.GradientStops.Item($n)
                $stops+=[ordered]@{position=[double]$stop.Position*100;rgb=[int]$stop.Color.RGB;transparency=[double]$stop.Transparency*100}
            }
            $cases+=[ordered]@{id=$id;preset=[int]$fill.PresetGradientType;angle=[double]$fill.GradientAngle;stops=$stops;chartXml=$chartXml}
        } finally {$probe.Close($false)}
    }
    [ordered]@{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;style=1;variant=1;cases=$cases} |
        ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'reference.json')
    Write-Output "Recorded $($cases.Count) native presets in $OutputFolder"
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
