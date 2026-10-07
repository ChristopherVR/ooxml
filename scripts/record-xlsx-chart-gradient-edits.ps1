param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-gradient-edits'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
function Capture($name) {
    $path=Join-Path $OutputFolder "$name.xlsx"; $book.SaveCopyAs($path)
    $probe=$excel.Workbooks.Open($path,0,$true)
    try {
        $fill=$probe.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Format.Fill
        if($name -like 'direction-*') {
            [void]$probe.Worksheets.Item(1).ChartObjects(1).Chart.Export((Join-Path $OutputFolder "$name.png"),'PNG')
        }
        $stops=@(); foreach($index in 1..$fill.GradientStops.Count){
            $stop=$fill.GradientStops.Item($index)
            $stops+=[ordered]@{position=[double]$stop.Position*100;transparency=[double]$stop.Transparency*100;brightness=[double]$stop.Color.Brightness*100;rgb=[int]$stop.Color.RGB}
        }
        $zip=[IO.Compression.ZipFile]::OpenRead($path)
        try{$reader=[IO.StreamReader]::new($zip.GetEntry('xl/charts/chart1.xml').Open());try{$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}}finally{$zip.Dispose()}
        return [ordered]@{name=$name;angle=[double]$fill.GradientAngle;stops=$stops;chartXml=$xml}
    } finally {$probe.Close($false)}
}
try {
    $excel=New-Object -ComObject Excel.Application; $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2=[string]'Month'; $sheet.Cells.Item(1,2).Value2=[string]'Sales'
    $sheet.Cells.Item(2,1).Value2=[string]'Jan'; $sheet.Cells.Item(2,2).Value2=[double]10
    $sheet.Cells.Item(3,1).Value2=[string]'Feb'; $sheet.Cells.Item(3,2).Value2=[double]20
    $chart=$sheet.ChartObjects().Add(200,20,480,300).Chart; $chart.ChartType=51
    $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $fill=$chart.SeriesCollection(1).Format.Fill; $fill.Solid()
    $fill.ForeColor.RGB=255; $fill.BackColor.RGB=16711680; $fill.TwoColorGradient(1,1)
    $cases=@(); $cases+=Capture 'initial'
    $fill.GradientAngle=[single]54
    $fill.GradientStops.Item(1).Position=[single]0.23
    $fill.GradientStops.Item(1).Transparency=[single]0.37
    $cases+=Capture 'edited'
    $fill.GradientStops.Insert(65280,[single]0.56,[single]0.13)
    $cases+=Capture 'inserted'
    $fill.GradientStops.Delete(2); $cases+=Capture 'removed'
    $minimumRejected=$false
    try{$fill.GradientStops.Delete(1)}catch{$minimumRejected=$true}
    foreach($value in @(-42,0,100,-100,37)) {
        $fill.GradientStops.Item(1).Color.Brightness=[single]($value/100)
        $cases+=Capture "brightness-$value"
    }
    $fill.GradientStops.Item(1).Color.RGB=16711680
    $cases+=Capture 'brightness-recolor'
    foreach($angle in @(0,45,90,135,180,225,270,315)) {
        $fill.GradientAngle=[single]$angle
        $cases+=Capture "direction-$angle"
    }
    [ordered]@{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;cases=$cases;minimumRejected=$minimumRejected} |
        ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'gradient-edits.json')
    Write-Output "Recorded gradient edits in $OutputFolder"
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
