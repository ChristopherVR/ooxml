param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-line-movement-'+[guid]::NewGuid().ToString('N'))),
 [switch]$DeleteAfterMove
)
# Capture endpoint translation without replacing native transform formulas.
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $directory){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
function Get-LineCells($shape){
 $result=[ordered]@{}
 foreach($name in @('BeginX','BeginY','EndX','EndY','Width','Height','PinX','PinY','LocPinX','LocPinY','Angle','FlipX','FlipY')){
  $cell=$shape.CellsU($name)
  $result[$name]=[ordered]@{formula=$cell.FormulaU;value=$cell.ResultIU}
 }
 return $result
}
try {
 $app.AlertResponse=7
 $document=$app.Documents.Add('')
 $page=$document.Pages.Item(1)
 $shapes=@()
 $cases=@()
 foreach($end in @(@(3,1.5),@(2.732050807568877,2.5),@(1,3.5),@(-1,1.5))){
  $shape=$page.DrawLine(1,1.5,$end[0],$end[1])
  if($shape.OneD -eq 0){throw 'Native probe is not a one-dimensional shape.'}
  $shapes+=,$shape
  $cases+=,[ordered]@{shapeId=[string]$shape.ID;oneD=$shape.OneD;before=(Get-LineCells $shape)}
 }
 $control=$null
 if($DeleteAfterMove){$control=$page.DrawRectangle(6,6,7,7)}
 $document.SaveAs((Join-Path $directory 'original.vsdx')) | Out-Null
 for($i=0;$i -lt $shapes.Length;$i++){
  $shape=$shapes[$i]
  $before=$cases[$i].before
  $dx=2.25;$dy=1.5
  foreach($name in @('BeginX','EndX')){$shape.CellsU($name).ResultIU=$before[$name].value+$dx}
  foreach($name in @('BeginY','EndY')){$shape.CellsU($name).ResultIU=$before[$name].value+$dy}
  $after=Get-LineCells $shape
  foreach($name in @('Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY')){
   if([Math]::Abs($before[$name].value-$after[$name].value) -gt 1e-12){throw "Translation changed $name."}
   if($before[$name].formula -ne $after[$name].formula){throw "Translation replaced the $name formula."}
  }
  foreach($axis in @('X','Y')){
   $delta=if($axis -eq 'X'){$dx}else{$dy}
   if([Math]::Abs($after['Pin'+$axis].value-$before['Pin'+$axis].value-$delta) -gt 1e-12){throw 'Native midpoint did not translate.'}
   if($before['Pin'+$axis].formula -ne $after['Pin'+$axis].formula){throw 'Native midpoint formula was replaced.'}
  }
  $cases[$i].Add('after',$after)
  $cases[$i].Add('delta',@($dx,$dy))
 }
 $document.SaveAs((Join-Path $directory 'moved.vsdx')) | Out-Null
 $evidence=[ordered]@{application='Microsoft Visio';version=$app.Version;cases=$cases}
 if($DeleteAfterMove){
  foreach($shape in $shapes){$shape.Delete()}
  if($page.Shapes.Count -ne 1 -or $page.Shapes.Item(1).ID -ne $control.ID){throw 'Deletion altered the control shape.'}
  $document.SaveAs((Join-Path $directory 'deleted.vsdx')) | Out-Null
  $evidence.Add('deletedShapeIds',@($cases | ForEach-Object shapeId))
  $evidence.Add('controlShapeId',[string]$control.ID)
 }
 $evidence | ConvertTo-Json -Depth 7 | Set-Content (Join-Path $directory 'evidence.json') -Encoding utf8
 Write-Output $directory
} finally {
 try {if($document){$document.Close()}} finally {$app.Quit()}
}
