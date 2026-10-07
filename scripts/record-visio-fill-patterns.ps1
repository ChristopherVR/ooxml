param(
 [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-fill-patterns-' + [guid]::NewGuid().ToString('N'))),
 [string]$Foreground = 'RGB(255,0,0)', [string]$Background = 'RGB(0,0,255)',
 [string]$ForegroundTransparency = '0%', [string]$BackgroundTransparency = '0%',
 [string]$Angle = '0 deg', [double]$DrawingScale = 1, [double]$PageScale = 1,
 [switch]$FlipX, [switch]$FlipY,
 [ValidateRange(0,8)][int]$GroupDepth = 0, [string]$GroupAngle = '0 deg',
 [switch]$GroupFlipX, [switch]$GroupFlipY,
 [ValidateRange(1,40)][int]$FirstPattern = 2, [ValidateRange(1,40)][int]$LastPattern = 24,
 [string]$GradientAngle = '',
 # -1 captures the modern direction matching each page's pattern number.
 [ValidateRange(-1,13)][int]$GradientDirection = 0,
 [ValidateSet('rectangle','ellipse','triangle','notched')][string]$ShapeKind = 'rectangle'
)
# Capture native pattern tiles and full-page exports from an owned application.
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'visio-native-gradient.ps1')
. (Join-Path $PSScriptRoot 'visio-native-shape.ps1')
if($FirstPattern -gt $LastPattern){throw 'FirstPattern must not exceed LastPattern.'}
if($GradientDirection -eq -1 -and $LastPattern -gt 13){throw 'Modern gradient directions must not exceed 13.'}
Add-Type -AssemblyName System.Drawing
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $directory){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
try {
 $app.AlertResponse=7
 $app.EventsEnabled=0
 $document=$app.Documents.Add('')
 $records=@()
 for($pattern=$FirstPattern;$pattern -le $LastPattern;$pattern++) {
  $page=if($pattern -eq $FirstPattern){$document.Pages.Item(1)}else{$document.Pages.Add()}
  $page.Name="Pattern-$pattern"
  $page.PageSheet.CellsU('DrawingScale').FormulaU="$DrawingScale in"
  $page.PageSheet.CellsU('PageScale').FormulaU="$PageScale in"
  $page.PageSheet.CellsU('PageWidth').FormulaU='4 in'
  $page.PageSheet.CellsU('PageHeight').FormulaU='3 in'
  $shape=New-VisioNativeFillShape $page $ShapeKind
  $shape.CellsU('Angle').FormulaU=$Angle
  $shape.CellsU('FlipX').FormulaU=if($FlipX){'1'}else{'0'}
  $shape.CellsU('FlipY').FormulaU=if($FlipY){'1'}else{'0'}
  $shape.CellsU('FillPattern').FormulaU=[string]$pattern
  $shape.CellsU('FillForegnd').FormulaU=$Foreground
  $shape.CellsU('FillBkgnd').FormulaU=$Background
  $shape.CellsU('FillForegndTrans').FormulaU=$ForegroundTransparency
  $shape.CellsU('FillBkgndTrans').FormulaU=$BackgroundTransparency
  $shape.CellsU('LinePattern').FormulaU='0'
  $modernGradient=$GradientAngle -or $GradientDirection -ne 0
  $direction=if($GradientDirection -eq -1){$pattern}else{$GradientDirection}
  if($modernGradient){
   $modernAngle=if($GradientAngle){$GradientAngle}else{'0 deg'}
   Set-VisioNativeFillGradient $shape $modernAngle $Foreground $Background $ForegroundTransparency $BackgroundTransparency $direction
  }
  $groupIds=@()
  for($depth=0;$depth -lt $GroupDepth;$depth++) {
   # An invisible sibling lets native Visio create a real group at each level.
   $sibling=$page.DrawRectangle(1,1,3,2)
   $sibling.CellsU('FillPattern').FormulaU='0'
   $sibling.CellsU('LinePattern').FormulaU='0'
   # visSelTypeAll=1; default iteration selects top-level shapes only.
   $selection=$page.CreateSelection(1)
   $group=$selection.Group()
   $group.CellsU('Angle').FormulaU=$GroupAngle
   $group.CellsU('FlipX').FormulaU=if($GroupFlipX){'1'}else{'0'}
   $group.CellsU('FlipY').FormulaU=if($GroupFlipY){'1'}else{'0'}
   $groupIds += [string]$group.ID
  }
  $svgPath=Join-Path $directory "pattern-$pattern.svg"
  $page.Export($svgPath)
  $page.Export((Join-Path $directory "pattern-$pattern.png"))
  $svg=[Xml.XmlDocument]::new();$svg.XmlResolver=$null
  $svg.LoadXml([IO.File]::ReadAllText($svgPath))
  $ns=[Xml.XmlNamespaceManager]::new($svg.NameTable)
  $ns.AddNamespace('s','http://www.w3.org/2000/svg')
  $ns.AddNamespace('x','http://www.w3.org/1999/xlink')
  $tile=$svg.SelectSingleNode('//s:pattern',$ns)
  $tileImage=if($tile){$tile.SelectSingleNode('s:image',$ns)}else{$null}
  if(-not $tileImage){
   if(-not $modernGradient -and $pattern -ge 2 -and $pattern -le 24){throw "Pattern $pattern has no native tile"}
   $records += [ordered]@{pattern=$pattern;pageId=[string]$page.ID;shapeId=[string]$shape.ID;groupIds=$groupIds}
   continue
  }
  $href=$tileImage.GetAttribute('href','http://www.w3.org/1999/xlink')
  $stream=[IO.MemoryStream]::new([Convert]::FromBase64String($href.Split(',')[1]))
  $bitmap=[Drawing.Bitmap]::new($stream)
  try {
   $pixels=@()
   for($y=0;$y -lt $bitmap.Height;$y++) {
    $row=@()
    for($x=0;$x -lt $bitmap.Width;$x++) {
     $pixel=$bitmap.GetPixel($x,$y)
     $row+=,@($pixel.R,$pixel.G,$pixel.B,$pixel.A)
    }
    $pixels+=,@($row)
   }
   $records += [ordered]@{pattern=$pattern;pageId=[string]$page.ID;shapeId=[string]$shape.ID;groupIds=$groupIds;widthPoints=[double]$tile.GetAttribute('width');heightPoints=[double]$tile.GetAttribute('height');pixelWidth=$bitmap.Width;pixelHeight=$bitmap.Height;pixels=$pixels}
  } finally {$bitmap.Dispose();$stream.Dispose()}
 }
 $document.SaveAs((Join-Path $directory 'fill-patterns.vsdx')) | Out-Null
 [ordered]@{application='Microsoft Visio';version=$app.Version;foreground=$Foreground;background=$Background;foregroundTransparency=$ForegroundTransparency;backgroundTransparency=$BackgroundTransparency;angle=$Angle;gradientAngle=$GradientAngle;gradientDirection=$GradientDirection;shapeKind=$ShapeKind;drawingScale=$DrawingScale;pageScale=$PageScale;flipX=[bool]$FlipX;flipY=[bool]$FlipY;groupDepth=$GroupDepth;groupAngle=$GroupAngle;groupFlipX=[bool]$GroupFlipX;groupFlipY=[bool]$GroupFlipY;firstPattern=$FirstPattern;lastPattern=$LastPattern;cases=$records} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'evidence.json') -Encoding utf8
} finally {if($document){$document.Saved=$true;$document.Close()};$app.Quit()}
Write-Output $directory
