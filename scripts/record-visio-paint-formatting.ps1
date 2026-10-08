# Fixed source paint references from an owned invisible Visio instance; no clipboard access.
param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-paint-formatting-'+[guid]::NewGuid().ToString('N'))),
 [string]$CoreOutputPath
)
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $directory|Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
function Read-Paint($shape) {
 $values=[ordered]@{}
 foreach($name in @('LinePattern','LineColorTrans','LineGradientEnabled','LineWeight','FillPattern','FillForegndTrans','FillBkgndTrans','FillGradientEnabled','PinX','PinY','Width','Height','Angle')) {
  $cell=$shape.CellsU($name);$values[$name]=[ordered]@{value=[double]$cell.ResultIU;formula=$cell.FormulaU}
 }
 [ordered]@{text=$shape.Text;values=$values;lineColor=$shape.CellsU('LineColor').ResultStrU(0);foreground=$shape.CellsU('FillForegnd').ResultStrU(0);background=$shape.CellsU('FillBkgnd').ResultStrU(0)}
}
function Apply-Paint($shape,$patch) {
 if($patch.fillColor) {
  $shape.CellsU('FillPattern').FormulaU=if($patch.fillColor -eq 'none'){'0'}else{'1'}
  $shape.CellsU('FillGradientEnabled').FormulaU='0'
  if($patch.fillColor -ne 'none'){$shape.CellsU('FillForegnd').FormulaU=Rgb $patch.fillColor;$shape.CellsU('FillForegndTrans').FormulaU='0%'}
 }
 if($null -ne $patch.fillPattern){$shape.CellsU('FillPattern').FormulaU=[string]$patch.fillPattern;$shape.CellsU('FillGradientEnabled').FormulaU='0'}
 if($patch.fillBackgroundColor){$shape.CellsU('FillBkgnd').FormulaU=Rgb $patch.fillBackgroundColor}
 if($null -ne $patch.fillTransparency){foreach($name in @('FillForegndTrans','FillBkgndTrans')){$shape.CellsU($name).FormulaU=([string]$patch.fillTransparency)+'%'}}
 if($patch.lineColor){$shape.CellsU('LineColor').FormulaU=Rgb $patch.lineColor;$shape.CellsU('LineColorTrans').FormulaU='0%';$shape.CellsU('LineGradientEnabled').FormulaU='0'}
 if($null -ne $patch.linePattern){$shape.CellsU('LinePattern').FormulaU=[string]$patch.linePattern}
 if($null -ne $patch.lineTransparency){$shape.CellsU('LineColorTrans').FormulaU=([string]$patch.lineTransparency)+'%'}
 if($null -ne $patch.lineWeight){$shape.CellsU('LineWeight').FormulaU=([string]$patch.lineWeight)+' pt'}
}
function Rgb([string]$color){'RGB('+((@(1,3,5)|ForEach-Object{[Convert]::ToInt32($color.Substring($_,2),16)}) -join ',')+')'}
try {
 $app.AlertResponse=7;$app.EventsEnabled=0
 if($CoreOutputPath) {
  $reference=Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw|ConvertFrom-Json
  $document=$app.Documents.OpenEx([IO.Path]::GetFullPath($CoreOutputPath),202)
  $accepted=@()
  for($index=0;$index -lt $reference.cases.Count;$index++) {
   $expected=$reference.cases[$index].native;$shape=$document.Pages.Item($index+1).Shapes.Item(1)
   $actual=Read-Paint $shape
   foreach($name in @('text','lineColor','foreground','background')){if($actual[$name] -cne $expected.$name){throw "Native reopen mismatch: $index $name"}}
   foreach($name in $actual.values.Keys){if([Math]::Abs($actual.values[$name].value-$expected.values.$name.value) -gt 1e-10){throw "Native reopen mismatch: $index $name"}}
   $accepted+=$actual
  }
  [ordered]@{version=$app.Version;accepted=$accepted}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
 } else {
  $document=$app.Documents.Add('')
  $patches=@(
   @{linePattern=0},@{linePattern=1},@{linePattern=2},@{linePattern=23},
   @{fillPattern=0},@{fillPattern=1},
   @{fillColor='#dc143c';fillPattern=2;fillBackgroundColor='#224466';fillTransparency=12.25},
   @{fillColor='#dc143c';fillPattern=24;fillBackgroundColor='#224466';fillTransparency=50},
   @{fillBackgroundColor='#224466'},
   @{lineTransparency=0.25},@{lineTransparency=12.75},@{lineTransparency=100},
   @{fillTransparency=0.25},@{fillTransparency=12.75},@{fillTransparency=100},
   @{fillColor='#ff0000';fillPattern=3;fillBackgroundColor='#0000ff';fillTransparency=50;lineColor='#123456';linePattern=2;lineTransparency=0.25;lineWeight=3},
   @{linePattern=1;setup='text-box'},@{lineTransparency=0;fillTransparency=0}
  )
  $cases=@()
  foreach($scale in @(1,2,0.5)){foreach($patch in $patches){
   $page=if($cases.Count -eq 0){$document.Pages.Item(1)}else{$document.Pages.Add()}
   $page.Name='Paint-'+[string]$cases.Count
   $page.PageSheet.CellsU('PageScale').FormulaU='1 in';$page.PageSheet.CellsU('DrawingScale').FormulaU=([string]$scale)+' in'
   $shape=$page.DrawRectangle(1,1,4,2);$shape.Text='Paint'
   if($patch.setup){$shape.CellsU('LinePattern').FormulaU='0';$shape.CellsU('FillPattern').FormulaU='0'}
   $command=[ordered]@{type='format-shape';pageId=[string]$page.ID;shapeId=[string]$shape.ID}
   foreach($key in $patch.Keys){if($key -ne 'setup'){$command[$key]=$patch[$key]}}
   $cases+=[ordered]@{drawingScale=$scale;command=$command}
  }}
  $document.SaveAs((Join-Path $directory 'original.vsdx'))|Out-Null
  for($index=0;$index -lt $cases.Count;$index++) {
   $shape=$document.Pages.Item($index+1).Shapes.Item(1)
   Apply-Paint $shape $cases[$index].command
   $cases[$index].native=Read-Paint $shape
  }
  $document.SaveAs((Join-Path $directory 'native.vsdx'))|Out-Null
  [ordered]@{version=$app.Version;cases=$cases}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
 }
}finally{if($document){$document.Saved=$true;$document.Close()};$app.Quit()}
Write-Output $directory
