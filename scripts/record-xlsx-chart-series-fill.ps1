param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-series-fill'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    foreach($col in 1..4){$sheet.Cells.Item(1,$col).Value2=[string]"C$col"}
    foreach($row in 2..5){
        $sheet.Cells.Item($row,1).Value2=[string]"M$row"
        foreach($col in 2..4){$sheet.Cells.Item($row,$col).Value2=[double]($row*$col)}
    }
    $chart=$sheet.ChartObjects().Add(260,20,480,300).Chart
    $chart.ChartType=57; $chart.SetSourceData($sheet.Range('A1:D5'),2)
    $first=$chart.SeriesCollection(1); $first.Points(1).Format.Fill.Solid()
    $first.Points(1).Format.Fill.ForeColor.RGB=65280
    $first.Format.Fill.Solid(); $first.Format.Fill.ForeColor.RGB=255
    $second=$chart.SeriesCollection(2)
    $second.Format.Fill.Solid(); $second.Format.Fill.ForeColor.ObjectThemeColor=5
    $second.Format.Fill.ForeColor.Brightness=0.4
    $third=$chart.SeriesCollection(3); $third.Points(1).Format.Fill.Solid()
    $third.Points(1).Format.Fill.ForeColor.RGB=65280; $third.Format.Fill.Visible=0
    $path=Join-Path $OutputFolder 'series-fills.xlsx'
    $book.SaveCopyAs($path); $book.Close($false); $book=$null
    $book=$excel.Workbooks.Open($path,0,$true)
    $native=$book.Worksheets.Item(1).ChartObjects(1).Chart
    $series=@()
    foreach($index in 1..3){
        $item=$native.SeriesCollection($index)
        $series+=@{index=$index-1;visible=[int]$item.Format.Fill.Visible;
            rgb=[int]$item.Format.Fill.ForeColor.RGB;
            brightness=[double]$item.Format.Fill.ForeColor.Brightness;
            firstPointVisible=[int]$item.Points(1).Format.Fill.Visible;
            firstPointRgb=[int]$item.Points(1).Format.Fill.ForeColor.RGB}
    }
    @{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;
        accent1Rgb=[int]$book.Theme.ThemeColorScheme.Colors(5).RGB;series=$series} |
        ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'series-fills.json')
    Write-Output "Recorded series fill behavior in $OutputFolder"
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
