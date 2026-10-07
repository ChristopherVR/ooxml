param(
 [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-fill-patterns-' + [guid]::NewGuid().ToString('N'))),
 [string]$Foreground = 'RGB(255,0,0)', [string]$Background = 'RGB(0,0,255)',
 [string]$ForegroundTransparency = '0%', [string]$BackgroundTransparency = '0%',
 [string]$Angle = '0 deg', [double]$DrawingScale = 1, [double]$PageScale = 1
)
# Capture native pattern tiles and full-page exports from an owned application.
$ErrorActionPreference='Stop'
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
 for($pattern=2;$pattern -le 24;$pattern++) {
  $page=if($pattern -eq 2){$document.Pages.Item(1)}else{$document.Pages.Add()}
  $page.Name="Pattern-$pattern"
  $page.PageSheet.CellsU('DrawingScale').FormulaU="$DrawingScale in"
  $page.PageSheet.CellsU('PageScale').FormulaU="$PageScale in"
  $page.PageSheet.CellsU('PageWidth').FormulaU='4 in'
  $page.PageSheet.CellsU('PageHeight').FormulaU='3 in'
  $shape=$page.DrawRectangle(1,1,3,2)
  $shape.CellsU('Angle').FormulaU=$Angle
  $shape.CellsU('FillPattern').FormulaU=[string]$pattern
  $shape.CellsU('FillForegnd').FormulaU=$Foreground
  $shape.CellsU('FillBkgnd').FormulaU=$Background
  $shape.CellsU('FillForegndTrans').FormulaU=$ForegroundTransparency
  $shape.CellsU('FillBkgndTrans').FormulaU=$BackgroundTransparency
  $shape.CellsU('LinePattern').FormulaU='0'
  $svgPath=Join-Path $directory "pattern-$pattern.svg"
  $page.Export($svgPath)
  $page.Export((Join-Path $directory "pattern-$pattern.png"))
  $svg=[Xml.XmlDocument]::new();$svg.XmlResolver=$null
  $svg.LoadXml([IO.File]::ReadAllText($svgPath))
  $ns=[Xml.XmlNamespaceManager]::new($svg.NameTable)
  $ns.AddNamespace('s','http://www.w3.org/2000/svg')
  $ns.AddNamespace('x','http://www.w3.org/1999/xlink')
  $tile=$svg.SelectSingleNode('//s:pattern',$ns)
  if(-not $tile){throw "Pattern $pattern has no native tile"}
  $href=$tile.SelectSingleNode('s:image',$ns).GetAttribute('href','http://www.w3.org/1999/xlink')
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
   $records += [ordered]@{pattern=$pattern;pageId=[string]$page.ID;shapeId=[string]$shape.ID;widthPoints=[double]$tile.GetAttribute('width');heightPoints=[double]$tile.GetAttribute('height');pixelWidth=$bitmap.Width;pixelHeight=$bitmap.Height;pixels=$pixels}
  } finally {$bitmap.Dispose();$stream.Dispose()}
 }
 $document.SaveAs((Join-Path $directory 'fill-patterns.vsdx')) | Out-Null
 [ordered]@{application='Microsoft Visio';version=$app.Version;foreground=$Foreground;background=$Background;foregroundTransparency=$ForegroundTransparency;backgroundTransparency=$BackgroundTransparency;angle=$Angle;drawingScale=$DrawingScale;pageScale=$PageScale;cases=$records} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'evidence.json') -Encoding utf8
} finally {if($document){$document.Saved=$true;$document.Close()};$app.Quit()}
Write-Output $directory
