param(
    [string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-chart-gradient-raster'),
    [ValidateSet('opaque','transparent','interior','three','three-transparent','crossed','coincident','path-corner','path-center','path-corner-transparent','path-center-transparent','path-center-circle','path-corner-circle','path-center-circle-transparent','path-corner-circle-transparent','path-center-shape','path-corner-shape','path-center-shape-transparent','path-corner-shape-transparent')]
    [string]$Profile = 'opaque',
    [ValidateSet('chart-area','series-column','series-bar')]
    [string]$Target = 'chart-area'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -ReferencedAssemblies System.Drawing.Common,System.Drawing.Primitives -TypeDefinition @'
using System.Drawing;
public static class NativeGradientBitmap {
    public static int[] Bounds(Bitmap bitmap) {
        int left=bitmap.Width, top=bitmap.Height, right=-1, bottom=-1;
        for(int y=0;y<bitmap.Height;y++) for(int x=0;x<bitmap.Width;x++) {
            // These red/white references exclude achromatic chart gridlines.
            var color=bitmap.GetPixel(x,y);
            if(color.A==0 || color.R<=color.G || color.R<=color.B) continue;
            left=System.Math.Min(left,x); top=System.Math.Min(top,y);
            right=System.Math.Max(right,x); bottom=System.Math.Max(bottom,y);
        }
        if(right<left) throw new System.Exception("Native series has no visible pixels");
        return new int[]{left,top,right-left+1,bottom-top+1};
    }
}
'@
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
        $directions = switch -Wildcard ($Profile) {
            'path-corner*' {@(1,2,3,4)}
            'path-center*' {@(1,2)}
            default {@(0,45,90,135,180,225,270,315,30,60)}
        }
        foreach($angle in $directions) {
            $object=$sheet.ChartObjects().Add(20,20,$size[0],$size[1]); $chart=$object.Chart
            $chart.ChartType=if($Target -eq 'series-bar'){57}else{51}; $chart.SetSourceData($sheet.Range('A1:A2'),2)
            $chart.HasTitle=$false; $chart.HasLegend=$false
            $chart.Axes(1).Delete(); $chart.Axes(2).Delete()
            $chart.PlotArea.Format.Fill.Visible=0; $chart.PlotArea.Format.Line.Visible=0
            $chart.SeriesCollection(1).Format.Fill.Visible=0
            $chart.SeriesCollection(1).Format.Line.Visible=0
            if($Target -eq 'chart-area') {
                $fill=$chart.ChartArea.Format.Fill
                $fillXPath='/c:chartSpace/c:spPr/a:gradFill'
            } else {
                $chart.ChartArea.Format.Fill.Visible=0
                $fill=$chart.SeriesCollection(1).Format.Fill
                $fillXPath='/c:chartSpace/c:chart/c:plotArea/c:barChart/c:ser/c:spPr/a:gradFill'
            }
            $fill.Visible=-1; $fill.Solid()
            $fill.ForeColor.RGB=255; $fill.TwoColorGradient(1,1)
            if($Profile -like 'path-corner*') {$fill.TwoColorGradient(5,$angle)}
            if($Profile -like 'path-center*') {$fill.TwoColorGradient(7,$angle)}
            $fill.GradientStops.Item(1).Color.RGB=255
            $fill.GradientStops.Item(2).Color.RGB=16777215
            switch($Profile) {
                'transparent' {
                    $fill.GradientStops.Item(1).Transparency=[single]0.37
                    $fill.GradientStops.Item(2).Transparency=[single]0.13
                }
                'interior' {
                    $fill.GradientStops.Item(1).Position=[single]0.23
                    $fill.GradientStops.Item(2).Position=[single]0.56
                }
                'three' {$fill.GradientStops.Insert(65280,[single]0.56,[single]0)}
                'three-transparent' {
                    $fill.GradientStops.Item(1).Transparency=[single]0.37
                    $fill.GradientStops.Insert(65280,[single]0.56,[single]0.13)
                }
                'crossed' {
                    $fill.GradientStops.Item(1).Position=[single]0.85
                    $fill.GradientStops.Item(2).Position=[single]0.56
                }
                'coincident' {
                    $fill.GradientStops.Item(1).Position=[single]0.5
                    $fill.GradientStops.Item(2).Position=[single]0.5
                }
            }
            if($Profile -like 'path-*-transparent') {
                $fill.GradientStops.Item(1).Transparency=[single]0.37
                $fill.GradientStops.Item(2).Transparency=[single]0.13
            }
            if($Profile -notlike 'path-*') {$fill.GradientAngle=[single]$angle}
            $chart.ChartArea.Format.Line.Visible=0
            $name="angle-$angle-$($size[0])x$($size[1])"
            if($Profile -ne 'opaque') {$name="$Profile-$name"}
            if($Target -ne 'chart-area') {$name="$Target-$name"}
            $path=Join-Path $OutputFolder "$name.xlsx"; $book.SaveCopyAs($path)
            $pathType = if($Profile -match '-(circle|shape)(-transparent)?$') {$Matches[1]} else {$null}
            if($pathType) {
                # COM only exposes legacy rectangular styles. Import the explicit OOXML
                # path, then let Excel save it again before measuring the native render.
                $package=[IO.Compression.ZipFile]::Open($path,[IO.Compression.ZipArchiveMode]::Update)
                try {
                    $entry=$package.GetEntry('xl/charts/chart1.xml')
                    $reader=[IO.StreamReader]::new($entry.Open())
                    try {[xml]$importXml=$reader.ReadToEnd()} finally {$reader.Dispose()}
                    $ns=[Xml.XmlNamespaceManager]::new($importXml.NameTable)
                    $ns.AddNamespace('c','http://schemas.openxmlformats.org/drawingml/2006/chart')
                    $ns.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
                    $importXml.SelectSingleNode("$fillXPath/a:path",$ns).SetAttribute('path',$pathType)
                    $entry.Delete(); $replacement=$package.CreateEntry('xl/charts/chart1.xml')
                    $writer=[IO.StreamWriter]::new($replacement.Open(),[Text.UTF8Encoding]::new($false))
                    try {$writer.Write($importXml.OuterXml)} finally {$writer.Dispose()}
                } finally {$package.Dispose()}
            }
            $zip=[IO.Compression.ZipFile]::OpenRead($path)
            try {
                $reader=[IO.StreamReader]::new($zip.GetEntry('xl/charts/chart1.xml').Open())
                try {[xml]$xml=$reader.ReadToEnd()} finally {$reader.Dispose()}
                $namespaces=[Xml.XmlNamespaceManager]::new($xml.NameTable)
                $namespaces.AddNamespace('c','http://schemas.openxmlformats.org/drawingml/2006/chart')
                $namespaces.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
                $fillXml=$xml.SelectSingleNode($fillXPath,$namespaces).OuterXml
            } finally {$zip.Dispose()}
            $probe=$excel.Workbooks.Open($path,0,$true)
            try {
                $native=$probe.Worksheets.Item(1).ChartObjects(1).Chart
                if($pathType) {
                    $saved=Join-Path $OutputFolder "$name-native.xlsx"; $probe.SaveCopyAs($saved)
                    $savedPackage=[IO.Compression.ZipFile]::OpenRead($saved)
                    try {
                        $reader=[IO.StreamReader]::new($savedPackage.GetEntry('xl/charts/chart1.xml').Open())
                        try {[xml]$savedXml=$reader.ReadToEnd()} finally {$reader.Dispose()}
                        $savedNs=[Xml.XmlNamespaceManager]::new($savedXml.NameTable)
                        $savedNs.AddNamespace('c','http://schemas.openxmlformats.org/drawingml/2006/chart')
                        $savedNs.AddNamespace('a','http://schemas.openxmlformats.org/drawingml/2006/main')
                        $fillXml=$savedXml.SelectSingleNode($fillXPath,$savedNs).OuterXml
                    } finally {$savedPackage.Dispose()}
                }
                $native.Parent.Activate(); $native.Refresh()
                $png=Join-Path $OutputFolder "$name.png"
                if(!$native.Export($png,'PNG')) {throw "Excel did not export $name"}
                $bitmap=[Drawing.Bitmap]::new($png)
                try {
                    $nativeFill=if($Target -eq 'chart-area'){$native.ChartArea.Format.Fill}else{$native.SeriesCollection(1).Format.Fill}
                    $nativeAngle=[double]$nativeFill.GradientAngle
                    $bounds=@(0,0,$bitmap.Width,$bitmap.Height)
                    if($Target -ne 'chart-area') {
                        # Gradient endpoints can contain white plateaus. Measure the same
                        # mark with a solid red fill instead of shrinking to colored pixels.
                        $nativeFill.Solid(); $nativeFill.ForeColor.RGB=255; $nativeFill.Transparency=[single]0
                        $native.Refresh(); $boundsPng=Join-Path $OutputFolder "$name-bounds.png"
                        if(!$native.Export($boundsPng,'PNG')) {throw "Excel did not export bounds for $name"}
                        $boundsBitmap=[Drawing.Bitmap]::new($boundsPng)
                        try {$bounds=[NativeGradientBitmap]::Bounds($boundsBitmap)} finally {$boundsBitmap.Dispose()}
                    }
                    $samples=@()
                    foreach($y in @(0.11,0.26,0.51,0.76,0.91)) {
                        foreach($x in @(0.11,0.26,0.51,0.76,0.91)) {
                            $px=$bounds[0]+[int][Math]::Floor($bounds[2]*$x)
                            $py=$bounds[1]+[int][Math]::Floor($bounds[3]*$y)
                            $color=$bitmap.GetPixel($px,$py)
                            $samples+=[ordered]@{x=$px;y=$py;rgb=@([int]$color.R,[int]$color.G,[int]$color.B);alpha=[int]$color.A}
                        }
                    }
                    $case=[ordered]@{name=$name;profile=$Profile;angle=$nativeAngle;fillXml=$fillXml;width=$bitmap.Width;height=$bitmap.Height;samples=$samples}
                    if($Target -ne 'chart-area') {$case['target']=$Target; $case['paintBounds']=[ordered]@{x=$bounds[0];y=$bounds[1];width=$bounds[2];height=$bounds[3]}}
                    $cases+=$case
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
