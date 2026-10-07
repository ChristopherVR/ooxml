param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-chart-gradient-raster'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2=[string]'Value'
    $sheet.Cells.Item(2,1).Value2=[double]1
    $cases=@()
    foreach($size in @(@(300,300),@(480,300))) {
        foreach($angle in @(0,45,90,135,180,225,270,315,30,60)) {
            $object=$sheet.ChartObjects().Add(20,20,$size[0],$size[1]); $chart=$object.Chart
            $chart.ChartType=51; $chart.SetSourceData($sheet.Range('A1:A2'),2)
            $chart.HasTitle=$false; $chart.HasLegend=$false
            $chart.Axes(1).Delete(); $chart.Axes(2).Delete()
            $chart.PlotArea.Format.Fill.Visible=0; $chart.PlotArea.Format.Line.Visible=0
            $chart.SeriesCollection(1).Format.Fill.Visible=0
            $chart.SeriesCollection(1).Format.Line.Visible=0
            $fill=$chart.ChartArea.Format.Fill; $fill.Solid()
            $fill.ForeColor.RGB=255; $fill.TwoColorGradient(1,1)
            $fill.GradientStops.Item(1).Color.RGB=255
            $fill.GradientStops.Item(2).Color.RGB=16777215
            $fill.GradientAngle=[single]$angle
            $chart.ChartArea.Format.Line.Visible=0
            $name="angle-$angle-$($size[0])x$($size[1])"
            $path=Join-Path $OutputFolder "$name.xlsx"; $book.SaveCopyAs($path)
            $zip=[IO.Compression.ZipFile]::OpenRead($path)
            try {
                $reader=[IO.StreamReader]::new($zip.GetEntry('xl/charts/chart1.xml').Open())
                try {[xml]$xml=$reader.ReadToEnd()} finally {$reader.Dispose()}
                $namespaces=[Xml.XmlNamespaceManager]::new($xml.NameTable)
                $namespaces.AddNamespace('c','http://schemas.openxmlformats.org/drawingml/2006/chart')
                $namespaces.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
                $fillXml=$xml.SelectSingleNode('/c:chartSpace/c:spPr/a:gradFill',$namespaces).OuterXml
            } finally {$zip.Dispose()}
            $probe=$excel.Workbooks.Open($path,0,$true)
            try {
                $native=$probe.Worksheets.Item(1).ChartObjects(1).Chart
                $native.Parent.Activate(); $native.Refresh()
                $png=Join-Path $OutputFolder "$name.png"
                if(!$native.Export($png,'PNG')) {throw "Excel did not export $name"}
                $bitmap=[Drawing.Bitmap]::new($png)
                try {
                    $samples=@()
                    foreach($y in @(0.11,0.26,0.51,0.76,0.91)) {
                        foreach($x in @(0.11,0.26,0.51,0.76,0.91)) {
                            $px=[int][Math]::Floor($bitmap.Width*$x)
                            $py=[int][Math]::Floor($bitmap.Height*$y)
                            $color=$bitmap.GetPixel($px,$py)
                            $samples+=[ordered]@{x=$px;y=$py;rgb=@([int]$color.R,[int]$color.G,[int]$color.B)}
                        }
                    }
                    $cases+=[ordered]@{name=$name;angle=[double]$native.ChartArea.Format.Fill.GradientAngle;fillXml=$fillXml;width=$bitmap.Width;height=$bitmap.Height;samples=$samples}
                } finally {$bitmap.Dispose()}
            } finally {$probe.Close($false)}
            $object.Delete()
        }
    }
    [ordered]@{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;cases=$cases} |
        ConvertTo-Json -Depth 8 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'reference.json')
    Write-Output "Recorded $($cases.Count) native chart gradients in $OutputFolder"
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
