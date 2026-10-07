param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-transparency'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2=[string]'Month'; $sheet.Cells.Item(1,2).Value2=[string]'Sales'
    $sheet.Cells.Item(2,1).Value2=[string]'Jan'; $sheet.Cells.Item(2,2).Value2=[double]10
    $sheet.Cells.Item(3,1).Value2=[string]'Feb'; $sheet.Cells.Item(3,2).Value2=[double]20
    $chart=$sheet.ChartObjects().Add(200,20,480,300).Chart
    $chart.ChartType=51; $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $series=$chart.SeriesCollection(1); $series.Format.Fill.Solid()
    $cases=@()
    foreach($percent in @(0,37,100)) {
        $series.Format.Fill.ForeColor.RGB=255
        $series.Format.Fill.Transparency=[double]($percent/100.0)
        $path=Join-Path $OutputFolder "opacity-$percent.xlsx"
        $book.SaveCopyAs($path)
        $probe=$excel.Workbooks.Open($path,0,$true)
        try {
            $native=$probe.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)
            $zip=[IO.Compression.ZipFile]::OpenRead($path)
            try {
                $reader=[IO.StreamReader]::new($zip.GetEntry('xl/charts/chart1.xml').Open())
                try{$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
            } finally {$zip.Dispose()}
            $cases+=@{percent=$percent;transparency=[double]$native.Format.Fill.Transparency;
                rgb=[int]$native.Format.Fill.ForeColor.RGB;chartXml=$xml}
        } finally {$probe.Close($false)}
    }
    $series.Format.Fill.Transparency=0.37; $series.Format.Fill.ForeColor.RGB=16711680
    $path=Join-Path $OutputFolder 'recolor.xlsx'; $book.SaveCopyAs($path)
    $probe=$excel.Workbooks.Open($path,0,$true)
    try{$colorChangeTransparency=[double]$probe.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Format.Fill.Transparency}finally{$probe.Close($false)}
    @{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;
        cases=$cases;colorChangeTransparency=$colorChangeTransparency} |
        ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'transparency.json')
    Write-Output "Recorded solid chart transparency in $OutputFolder"
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
