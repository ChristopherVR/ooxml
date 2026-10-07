param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-gradient-raster-'+[guid]::NewGuid().ToString('N'))),
 [ValidateSet('rectangle','ellipse','triangle','notched','pentagon','chevron','ushape','star')]
 [string[]]$PathShapes=@('rectangle','ellipse','triangle','notched'),
 [ValidateRange(-180,180)][double]$ShapeAngle=0,
 [ValidateRange(0,13)][int]$FirstDirection=0,
 [ValidateRange(0,13)][int]$LastDirection=13,
 [ValidateSet('Fill','Line')][string]$Paint='Fill',
 [string]$GradientAngle='0 deg'
)
# Compare the real raster engine, rather than assuming native SVG is a paint oracle.
$ErrorActionPreference='Stop'
if($FirstDirection -gt $LastDirection){throw 'FirstDirection must not exceed LastDirection.'}
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
   foreach($direction in $FirstDirection..$LastDirection){
    $kinds=if($direction -eq 13){$PathShapes}else{@('rectangle')}
    foreach($kind in $kinds){
     $name="direction-$direction-$kind-stops-$stopCount-alpha-$alpha"
     if($Paint -eq 'Line'){$name='line-'+$name}
     if($ShapeAngle -ne 0){$name+='-angle-'+$ShapeAngle.ToString([cultureinfo]::InvariantCulture)}
     $shape=New-VisioNativeFillShape $page $kind
     $shape.CellsU('Angle').FormulaU=$ShapeAngle.ToString([cultureinfo]::InvariantCulture)+' deg'
     $shape.CellsU('FillPattern').FormulaU='1'
     $shape.CellsU('LinePattern').FormulaU='0'
     if($Paint -eq 'Line'){
      $shape.CellsU('FillPattern').FormulaU='0'
      $shape.CellsU('LinePattern').FormulaU='1'
      $shape.CellsU('LineWeight').FormulaU='0.1 in'
     }
     $front=if($alpha){'20%'}else{'0%'}
     $back=if($alpha){'50%'}else{'0%'}
     Set-VisioNativeGradient $shape $Paint $GradientAngle 'RGB(255,0,0)' 'RGB(0,0,255)' $front $back $direction
     $gradientSection=if($Paint -eq 'Line'){248}else{249}
     if($stopCount -eq 3){
      while($shape.RowCount($gradientSection) -lt 3){$shape.AddRow($gradientSection,-1,0) | Out-Null}
      $shape.CellsSRC($gradientSection,1,0).FormulaU='RGB(0,255,0)'
      $shape.CellsSRC($gradientSection,1,1).FormulaU=if($alpha){'35%'}else{'0%'}
      $shape.CellsSRC($gradientSection,1,2).FormulaU='50%'
      $shape.CellsSRC($gradientSection,2,0).FormulaU='RGB(0,0,255)'
      $shape.CellsSRC($gradientSection,2,1).FormulaU=$back
      $shape.CellsSRC($gradientSection,2,2).FormulaU='100%'
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
      $outline=@()
      if($kind -ne 'ellipse'){
       $points=[double[]](Get-VisioNativeFillPoints $kind)
       for($i=0;$i -lt $points.Length-2;$i+=2){$outline+=,@((($points[$i]-1)/2),($points[$i+1]-1))}
      }
      # Native page-space pose and extents register rotated fills without inferring them from our parser.
      # https://learn.microsoft.com/en-us/office/vba/api/visio.shape.boundingbox
      $left=0.0;$bottom=0.0;$right=0.0;$top=0.0
      $shape.BoundingBox(8196,[ref]$left,[ref]$bottom,[ref]$right,[ref]$top)
      $nativeExtents=@($left,$bottom,$right,$top)
      if($left -gt $right -or $bottom -gt $top){throw 'Native geometry export bounds are empty.'}
      $ox=0.0;$oy=0.0;$xx=0.0;$xy=0.0;$yx=0.0;$yy=0.0
      $shape.XYToPage(0,0,[ref]$ox,[ref]$oy)
      $shape.XYToPage(1,0,[ref]$xx,[ref]$xy)
      $shape.XYToPage(0,1,[ref]$yx,[ref]$yy)
      $nativeTransform=@(($xx-$ox),($xy-$oy),($yx-$ox),($yy-$oy),$ox,$oy)
      $cases+=,@{name=$name;paint=$Paint;direction=$direction;kind=$kind;outline=$outline;angle=$ShapeAngle;nativeExtents=$nativeExtents;nativeLineWidth=$shape.CellsU('LineWeight').ResultIU;nativeTransform=$nativeTransform;stopCount=$stopCount;alpha=$alpha;shapeId=[string]$shape.ID;width=$bitmap.Width;height=$bitmap.Height;samples=$samples}
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
