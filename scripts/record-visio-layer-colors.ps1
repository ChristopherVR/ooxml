param([string]$OutputDirectory = (Join-Path $env:TEMP ('visio-layer-colors-' + [guid]::NewGuid().ToString('N'))), [switch]$VerifyCore)
# Owned native Visio SVG/PNG oracle; no user application instance is accessed.
Add-Type -AssemblyName System.Drawing
function Read-Paint($directory, $name) {
 $svg=[Xml.XmlDocument]::new(); $svg.XmlResolver=$null
 $svg.LoadXml([IO.File]::ReadAllText((Join-Path $directory "$name.svg")))
 $namespaces=[Xml.XmlNamespaceManager]::new($svg.NameTable)
 $namespaces.AddNamespace('s','http://www.w3.org/2000/svg')
 $classes=@{}
 foreach($match in [regex]::Matches($svg.SelectSingleNode('//s:style',$namespaces).InnerText,'\.(st\d+)\s*\{([^}]*)\}')) {
  $style=@{}
  foreach($entry in $match.Groups[2].Value.Split(';')) {
   $pair=$entry.Split(':',2); if($pair.Length -eq 2){$style[$pair[0].Trim()]=$pair[1].Trim()}
  }
  $classes[$match.Groups[1].Value]=$style
 }
 $paint=$classes[$svg.SelectSingleNode('//s:rect',$namespaces).GetAttribute('class')]
 $text=$classes[$svg.SelectSingleNode('//s:text',$namespaces).GetAttribute('class')]
 $gradient=$svg.SelectSingleNode('//s:linearGradient',$namespaces)
 $stops=@();$angle=$null
 if($gradient) {
  $angle=[double][regex]::Match($gradient.GetAttribute('gradientTransform'),'rotate\(([^ ]+)').Groups[1].Value
  foreach($stop in $gradient.SelectNodes('s:stop',$namespaces)) {
   $stops += [ordered]@{offset=[double]$stop.GetAttribute('offset');color=$stop.GetAttribute('stop-color');opacity=[double]$stop.GetAttribute('stop-opacity')}
  }
 }
 $bitmap=[Drawing.Bitmap]::new((Join-Path $directory "$name.png"))
 try { $pixel=$bitmap.GetPixel([int]($bitmap.Width/2),1); $strokePixel=@($pixel.R,$pixel.G,$pixel.B,$pixel.A) } finally { $bitmap.Dispose() }
 [ordered]@{fill=$paint.fill;fillOpacity=$(if($paint.ContainsKey('fill-opacity')){[double]$paint['fill-opacity']}else{1})
  stroke=$paint.stroke;strokeOpacity=$(if($paint.ContainsKey('stroke-opacity')){[double]$paint['stroke-opacity']}else{1})
  text=$text.fill;textOpacity=$(if($text.ContainsKey('opacity')){[double]$text.opacity}else{1})
  gradientAngle=$angle;gradientStops=$stops;strokePixel=$strokePixel}
}
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'visio-native-gradient.ps1')
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if (-not $VerifyCore -and (Test-Path -LiteralPath $directory) -and @(Get-ChildItem -LiteralPath $directory).Count) { throw 'Use a fresh output directory to avoid native overwrite dialogs.' }
if(-not (Test-Path -LiteralPath $directory)){New-Item -ItemType Directory -Path $directory | Out-Null}
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
try {
 $app.AlertResponse=7
 $app.EventsEnabled=0
 if($VerifyCore) {
  Write-Output "Opening core drawing with owned Visio instance $($app.ProcessID)"
  $original=Get-Content (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
  Write-Output 'Opening native original as a hidden drawing'
  $document=$app.Documents.OpenEx((Join-Path $directory 'layer-colors.vsdx'),202)
  Write-Output 'Native original reopened'
  $document.Saved=$true; $document.Close(); $document=$null
  $document=$app.Documents.OpenEx((Join-Path $directory 'core-layer-colors.vsdx'),202)
  Write-Output 'Core drawing opened'
  if($document.Pages.Count -ne $original.cases.Count){throw 'Core page count changed'}
  if($document.Pages.Item(1).Shapes.Item(1).CellsU('PinX').ResultIU -ne 2.25){throw 'Core move not accepted'}
  $accepted=@()
  for($index=0;$index -lt $original.cases.Count;$index++) {
   $case=$original.cases[$index]; $page=$document.Pages.Item($index+1); $name='core-'+$case.name
   if([string]$page.ID -ne $case.pageId){throw 'Core page ID changed'}
   $page.Export((Join-Path $directory "$name.svg")); $page.Export((Join-Path $directory "$name.png"))
   $paint=Read-Paint $directory $name
   foreach($key in @('fill','fillOpacity','stroke','strokeOpacity','text','textOpacity','gradientAngle')) {
    if($key -eq 'fill' -and $case.paint.fill.StartsWith('url(') -and $paint.fill.StartsWith('url(')){continue}
    if($paint[$key] -ne $case.paint.$key){throw "Core $($case.name) $key changed"}
   }
   if(($paint.gradientStops|ConvertTo-Json -Depth 6 -Compress) -ne ($case.paint.gradientStops|ConvertTo-Json -Depth 6 -Compress)){throw 'Core gradient colors changed'}
   $accepted += [ordered]@{name=$case.name;pageId=[string]$page.ID;paint=$paint}
  }
  $saved=Join-Path $directory 'core-layer-colors-native-reopened.vsdx'
  if(Test-Path -LiteralPath $saved){throw 'Use fresh core verification output names'}
  $document.SaveAs($saved) | Out-Null
  [ordered]@{application='Microsoft Visio';version=$app.Version;cases=$accepted} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'core-evidence.json') -Encoding utf8
  Write-Output 'Core drawing accepted, exported and saved by native Visio'
  return
 }
 $document=$app.Documents.Add('')
 $cases=@('opaque','partial','transparent','none','multiple','mixed','source-alpha','no-fill','no-line','gradient','hatch','uncolored-alpha','modern-gradient','modern-gradient-partial','fractional','linear26','linear27','linear28','linear29','linear30','linear-partial')
 $record=@()
 foreach($name in $cases) {
  if($name -eq 'opaque') { $page=$document.Pages.Item(1) } else { $page=$document.Pages.Add() }
  $page.Name="Layer-$name"
  $page.PageSheet.CellsU('PageWidth').FormulaU='4 in'
  $page.PageSheet.CellsU('PageHeight').FormulaU='3 in'
  $shape=$page.DrawRectangle(1,1,3,2)
  $shape.Text='Layer color'
  $shape.CellsU('FillForegnd').FormulaU='RGB(0,255,0)'
  $shape.CellsU('LineColor').FormulaU='RGB(0,0,255)'
  $shape.CellsU('Char.Color').FormulaU='RGB(0,255,255)'
  $shape.CellsU('LineWeight').FormulaU='0.03 in'
  $layer=$page.Layers.Add('Red')
  $layer.CellsC(2).FormulaU=$(if($name -eq 'none') {'255'} else {'RGB(255,0,0)'})
  $layer.CellsC(11).FormulaU=$(if($name -eq 'partial' -or $name -eq 'source-alpha') {'40%'} elseif($name -eq 'transparent') {'100%'} else {'0%'})
  $layer.Add($shape,1)
  if($name -eq 'multiple' -or $name -eq 'mixed') {
   $other=$page.Layers.Add('Other')
   $other.CellsC(2).FormulaU=$(if($name -eq 'multiple') {'RGB(0,0,255)'} else {'255'})
   $other.Add($shape,1)
  }
  if($name -eq 'source-alpha') {
   $shape.CellsU('FillForegndTrans').FormulaU='30%'
   $shape.CellsU('LineColorTrans').FormulaU='20%'
   $shape.CellsU('Char.ColorTrans').FormulaU='50%'
  }
  if($name -eq 'no-fill') { $shape.CellsU('FillPattern').FormulaU='0' }
  if($name -eq 'fractional') { $layer.CellsC(11).FormulaU='41.3%' }
  if($name -match '^linear(26|27|28|29|30)$') { $shape.CellsU('FillPattern').FormulaU=$Matches[1] }
  if($name -eq 'linear-partial') { $shape.CellsU('FillPattern').FormulaU='25'; $layer.CellsC(11).FormulaU='40%' }
  if($name -eq 'no-line') { $shape.CellsU('LinePattern').FormulaU='0' }
  if($name -eq 'gradient') { $shape.CellsU('FillPattern').FormulaU='25' }
  if($name -eq 'hatch') { $shape.CellsU('FillPattern').FormulaU='2' }
  if($name -eq 'uncolored-alpha') { $layer.CellsC(2).FormulaU='255'; $shape.CellsU('Char.ColorTrans').FormulaU='50%' }
  if($name -like 'modern-gradient*') {
   Set-VisioNativeLinearGradient $shape '0 deg' 'RGB(0,255,0)' 'RGB(0,0,255)' '20%' '50%'
   if($name -eq 'modern-gradient-partial') { $layer.CellsC(11).FormulaU='40%' }
  }
  $page.Export((Join-Path $directory "$name.svg"))
  $page.Export((Join-Path $directory "$name.png"))
  $record += [ordered]@{name=$name;pageId=[string]$page.ID;shapeId=[string]$shape.ID;members=$shape.CellsU('LayerMember').ResultStrU(0);layerColor=$layer.CellsC(2).FormulaU;layerTransparency=$layer.CellsC(11).ResultIU;paint=(Read-Paint $directory $name)}
 }
 $document.SaveAs((Join-Path $directory 'layer-colors.vsdx')) | Out-Null
 [ordered]@{application='Microsoft Visio';version=$app.Version;cases=$record} | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $directory 'evidence.json') -Encoding utf8
} finally { if($document){$document.Saved=$true;$document.Close()};$app.Quit() }
Write-Output $directory
