param([string]$OutputDirectory=(Join-Path $env:TEMP ('visio-gradient-raster-'+[guid]::NewGuid().ToString('N'))))
# Compare the real raster engine, rather than assuming native SVG is a paint oracle.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'visio-native-shape.ps1')
. (Join-Path $PSScriptRoot 'visio-native-gradient.ps1')
Add-Type -AssemblyName System.Drawing
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $directory){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$brush=[Drawing.Drawing2D.LinearGradientBrush]::new([Drawing.Point]::new(0,0),[Drawing.Point]::new(288,0),[Drawing.Color]::Red,[Drawing.Color]::Blue)
try {
 $brush.SetSigmaBellShape(1,1)
 [ordered]@{positions=$brush.Blend.Positions;factors=$brush.Blend.Factors} | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $directory 'sigma-blend.json') -Encoding utf8
} finally {$brush.Dispose()}
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
$size=0;$width=0.0;$height=0.0;$sizeUnits=0
$resolution=0;$resolutionWidth=0.0;$resolutionHeight=0.0;$resolutionUnits=0
$settingsCaptured=$false
try {
 $app.Settings.GetRasterExportSize([ref]$size,[ref]$width,[ref]$height,[ref]$sizeUnits)
 $app.Settings.GetRasterExportResolution([ref]$resolution,[ref]$resolutionWidth,[ref]$resolutionHeight,[ref]$resolutionUnits)
 $settingsCaptured=$true
 $previous=[ordered]@{size=$size;width=$width;height=$height;sizeUnits=$sizeUnits;resolution=$resolution;resolutionWidth=$resolutionWidth;resolutionHeight=$resolutionHeight;resolutionUnits=$resolutionUnits}
 $previous | ConvertTo-Json | Set-Content (Join-Path $directory 'previous-settings.json') -Encoding utf8
 # Capture the same 2-by-1-inch shape at exactly 144 dpi.
 $app.Settings.SetRasterExportSize(3,288,144,0)
 $app.Settings.SetRasterExportResolution(3,144,144,0)
 $app.AlertResponse=7
 $document=$app.Documents.Add('')
 $page=$document.Pages.Item(1)
 $cases=@()
 foreach($stopCount in @(2,3)){
  foreach($alpha in @($false,$true)){
   foreach($direction in 0..13){
    $kinds=if($direction -eq 13){@('rectangle','ellipse','triangle','notched')}else{@('rectangle')}
    foreach($kind in $kinds){
     $name="direction-$direction-$kind-stops-$stopCount-alpha-$alpha"
     $shape=New-VisioNativeFillShape $page $kind
     $shape.CellsU('FillPattern').FormulaU='1'
     $shape.CellsU('LinePattern').FormulaU='0'
     $front=if($alpha){'20%'}else{'0%'}
     $back=if($alpha){'50%'}else{'0%'}
     Set-VisioNativeFillGradient $shape '0 deg' 'RGB(255,0,0)' 'RGB(0,0,255)' $front $back $direction
     if($stopCount -eq 3){
      while($shape.RowCount(249) -lt 3){$shape.AddRow(249,-1,0) | Out-Null}
      $shape.CellsSRC(249,1,0).FormulaU='RGB(0,255,0)'
      $shape.CellsSRC(249,1,1).FormulaU=if($alpha){'35%'}else{'0%'}
      $shape.CellsSRC(249,1,2).FormulaU='50%'
      $shape.CellsSRC(249,2,0).FormulaU='RGB(0,0,255)'
      $shape.CellsSRC(249,2,1).FormulaU=$back
      $shape.CellsSRC(249,2,2).FormulaU='100%'
     }
     $shape.Export((Join-Path $directory ($name+'.svg')))
     $shape.Export((Join-Path $directory ($name+'.png')))
     $bitmap=[Drawing.Bitmap]::FromFile((Join-Path $directory ($name+'.png')))
     try {
      $samples=@()
      foreach($point in @(@(36,72),@(72,72),@(144,72),@(216,72),@(72,36),@(144,36),@(144,96))){
       $color=$bitmap.GetPixel($point[0],$point[1])
       $samples+=,@{x=$point[0];y=$point[1];rgba=@($color.R,$color.G,$color.B,$color.A)}
      }
      $cases+=,@{name=$name;direction=$direction;kind=$kind;stopCount=$stopCount;alpha=$alpha;shapeId=[string]$shape.ID;width=$bitmap.Width;height=$bitmap.Height;samples=$samples}
     } finally {$bitmap.Dispose()}
    }
   }
  }
 }
 $document.SaveAs((Join-Path $directory 'gradient-raster.vsdx')) | Out-Null
 [ordered]@{application='Microsoft Visio';version=$app.Version;dpi=144;cases=$cases} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'evidence.json') -Encoding utf8
} finally {
 try {
  if($settingsCaptured){
   try {$app.Settings.SetRasterExportSize($size,$width,$height,$sizeUnits)}
   finally {$app.Settings.SetRasterExportResolution($resolution,$resolutionWidth,$resolutionHeight,$resolutionUnits)}
  }
 } finally {
  try {if($document){$document.Saved=$true;$document.Close()}}
  finally {$app.Quit()}
 }
}
Write-Output $directory
