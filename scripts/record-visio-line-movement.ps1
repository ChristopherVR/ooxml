param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-line-movement-'+[guid]::NewGuid().ToString('N'))),
 [switch]$DeleteAfterMove,
 [ValidateRange(0,1000000)][double]$ResizeWidth=0,
 [ValidateSet('None','Begin','End')][string]$MoveEndpoint='None',
 [ValidateRange(0.000001,1000000)][double]$DrawingScale=1,
 [ValidateRange(0.000001,1000000)][double]$PageScale=1,
 [switch]$GridAligned,
 [switch]$IncludeRectangle,
 [switch]$IncludeEllipse,
 [ValidateRange(-360000,360000)][double]$RotationDegrees=0,
 [switch]$OffCentrePin,
 [ValidateSet('None','Left','Right')][string]$QuarterTurn='None',
 [ValidateSet('None','Horizontal','Vertical')][string]$Flip='None',
 [ValidateSet('None','LockRotate','GuardAngle')][string]$FlipProtection='None',
 [switch]$CustomDefaults
)
# Capture endpoint translation without replacing native transform formulas.
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $directory){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
function Get-ShapeCells($shape,[string[]]$names){
 $result=[ordered]@{}
 foreach($name in $names){$cell=$shape.CellsU($name);$result[$name]=[ordered]@{formula=$cell.FormulaU;value=$cell.ResultIU}}
 return $result
}
function Get-LineCells($shape){return Get-ShapeCells $shape @('BeginX','BeginY','EndX','EndY','Width','Height','PinX','PinY','LocPinX','LocPinY','Angle','FlipX','FlipY')}
function Get-LineTransform($shape){
 $x0=0.0;$y0=0.0;$xx=0.0;$yx=0.0;$xy=0.0;$yy=0.0
 $shape.XYToPage(0,0,[ref]$x0,[ref]$y0)
 $shape.XYToPage(1,0,[ref]$xx,[ref]$yx)
 $shape.XYToPage(0,1,[ref]$xy,[ref]$yy)
 return @(($xx-$x0),($yx-$y0),($xy-$x0),($yy-$y0),$x0,$y0)
}
try {
 $app.AlertResponse=7
 $document=$app.Documents.Add('')
 if($CustomDefaults){
  $base=$document.DefaultStyle
  $line=$document.Styles.Add('Parity line default',$base,1,1,1)
  $fill=$document.Styles.Add('Parity fill default',$base,1,1,1)
  $text=$document.Styles.Add('Parity text default',$base,1,1,1)
  $line.CellsU('LineColor').FormulaU='RGB(24,96,168)'
  $line.CellsU('LineWeight').ResultIU=0.05
  $fill.CellsU('FillForegnd').FormulaU='RGB(240,176,80)'
  $fill.CellsU('FillPattern').ResultIU=1
  $text.CellsU('Char.Color').FormulaU='RGB(128,32,96)'
  $document.DefaultLineStyle=$line.Name
  $document.DefaultFillStyle=$fill.Name
  $document.DefaultTextStyle=$text.Name
 }
 $page=$document.Pages.Item(1)
 $page.PageSheet.CellsU('DrawingScale').ResultIU=$DrawingScale
 $page.PageSheet.CellsU('PageScale').ResultIU=$PageScale
 $shapes=@()
 $cases=@()
 # GridAligned references use identical physical-page coordinates across scales.
 [double]$coordinates=if($GridAligned){$DrawingScale/$PageScale}else{1.0}
 $diagonal=if($GridAligned){@(3,3.5)}else{@(2.732050807568877,2.5)}
 foreach($end in @(@(3,1.5),$diagonal,@(1,3.5),@(-1,1.5))){
  $shape=$page.DrawLine($coordinates,1.5*$coordinates,$end[0]*$coordinates,$end[1]*$coordinates)
  if($shape.OneD -eq 0){throw 'Native probe is not a one-dimensional shape.'}
  $shapes+=,$shape
  $cases+=,[ordered]@{shapeId=[string]$shape.ID;oneD=$shape.OneD;before=(Get-LineCells $shape);beforeTransform=(Get-LineTransform $shape)}
 }
 $rectangleEvidence=$null
 if($IncludeRectangle){
  $rectangle=$page.DrawRectangle(5.5*$coordinates,2*$coordinates,7.5*$coordinates,3*$coordinates)
  $rectangleCells=Get-ShapeCells $rectangle @('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY','ResizeMode','QuickStyleLineMatrix','QuickStyleFillMatrix','QuickStyleEffectsMatrix','QuickStyleFontMatrix')
  $rectangleEvidence=[ordered]@{shapeId=[string]$rectangle.ID;cells=$rectangleCells;transform=(Get-LineTransform $rectangle)}
 }
 $ellipseEvidence=$null
 if($IncludeEllipse){
  $ellipse=$page.DrawOval(5.5*$coordinates,5*$coordinates,7.5*$coordinates,6*$coordinates)
  $ellipseNames=@('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY','ResizeMode')
  $ellipseCells=Get-ShapeCells $ellipse $ellipseNames
  $ellipseEvidence=[ordered]@{shapeId=[string]$ellipse.ID;cells=$ellipseCells;transform=(Get-LineTransform $ellipse)}
 }
 $control=$null
 if($DeleteAfterMove){$control=$page.DrawRectangle(6,6,7,7)}
 $document.SaveAs((Join-Path $directory 'original.vsdx')) | Out-Null
 $page.Export((Join-Path $directory 'original-page.svg'))
 for($i=0;$i -lt $shapes.Length;$i++){
  $shape=$shapes[$i]
  $before=$cases[$i].before
  $dx=2.25*$coordinates;$dy=1.5*$coordinates
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
  $cases[$i].Add('afterTransform',(Get-LineTransform $shape))
  $cases[$i].Add('delta',@($dx,$dy))
 }
 $document.SaveAs((Join-Path $directory 'moved.vsdx')) | Out-Null
 $page.Export((Join-Path $directory 'moved-page.svg'))
 if($ResizeWidth -gt 0){
  for($i=0;$i -lt $shapes.Length;$i++){
   $shapes[$i].CellsU('Width').ResultIU=$ResizeWidth
   $cases[$i].Add('resized',(Get-LineCells $shapes[$i]))
   $cases[$i].Add('resizedTransform',(Get-LineTransform $shapes[$i]))
  }
  $document.SaveAs((Join-Path $directory 'resized.vsdx')) | Out-Null
 }
 if($MoveEndpoint -ne 'None'){
  for($i=0;$i -lt $shapes.Length;$i++){
   $shape=$shapes[$i]
   $cases[$i].Add('endpointBefore',(Get-LineCells $shape))
   $shape.CellsU($MoveEndpoint+'X').ResultIU+=0.75
   $shape.CellsU($MoveEndpoint+'Y').ResultIU-=0.5
   $cases[$i].Add('endpointAfter',(Get-LineCells $shape))
   $cases[$i].Add('endpointTransform',(Get-LineTransform $shape))
   $cases[$i].Add('endpoint',$MoveEndpoint)
  }
  $document.SaveAs((Join-Path $directory 'endpoint.vsdx')) | Out-Null
  $page.Export((Join-Path $directory 'endpoint-page.svg'))
 }
 if($IncludeEllipse){
  $ellipse.CellsU('Width').ResultIU=3*$coordinates
  $ellipse.CellsU('Height').ResultIU=2*$coordinates
  $ellipse.CellsU('PinX').ResultIU+=0.5*$coordinates
  $ellipse.CellsU('PinY').ResultIU+=0.5*$coordinates
  $ellipseEvidence.Add('editedCells',(Get-ShapeCells $ellipse $ellipseNames))
  $ellipseEvidence.Add('editedTransform',(Get-LineTransform $ellipse))
  $document.SaveAs((Join-Path $directory 'ellipse-edited.vsdx')) | Out-Null
  $page.Export((Join-Path $directory 'ellipse-edited-page.svg'))
 }
 $evidence=[ordered]@{application='Microsoft Visio';version=$app.Version;cases=$cases;drawingScale=$page.PageSheet.CellsU('DrawingScale').ResultIU;pageScale=$page.PageSheet.CellsU('PageScale').ResultIU}
 if($rectangleEvidence){$evidence.Add('rectangle',$rectangleEvidence)}
 if($ellipseEvidence){$evidence.Add('ellipse',$ellipseEvidence)}
 if($RotationDegrees -ne 0){
  if($OffCentrePin){
   foreach($shape in @($rectangle,$ellipse)){
    if($shape){$shape.CellsU('LocPinX').ResultIU=$shape.CellsU('Width').ResultIU*0.25;$shape.CellsU('LocPinY').ResultIU=$shape.CellsU('Height').ResultIU*0.75}
   }
  }
  $document.SaveAs((Join-Path $directory 'rotation-source.vsdx')) | Out-Null
  $rotated=[ordered]@{}
  foreach($entry in @(@('rectangle',$rectangle),@('ellipse',$ellipse))){
   if($entry[1]){
    $entry[1].CellsU('Angle').ResultIU=$RotationDegrees*[Math]::PI/180
    $rotated[$entry[0]]=[ordered]@{shapeId=[string]$entry[1].ID;cells=(Get-ShapeCells $entry[1] @('Width','Height','PinX','PinY','LocPinX','LocPinY','Angle'));transform=(Get-LineTransform $entry[1])}
   }
  }
  $document.SaveAs((Join-Path $directory 'rotated.vsdx')) | Out-Null
  $page.Export((Join-Path $directory 'rotated-page.svg'))
  $evidence.Add('rotated',$rotated)
 }
 if($QuarterTurn -ne 'None'){
  $document.SaveAs((Join-Path $directory 'quarter-source.vsdx')) | Out-Null
  $turned=[ordered]@{}
  $degrees=if($QuarterTurn -eq 'Left'){90}else{-90}
  foreach($entry in @(@('rectangle',$rectangle),@('ellipse',$ellipse))){
   if($entry[1]){
    $selection=$page.CreateSelection(2,256,$entry[1])
    $selection.Rotate($degrees,'deg',$false,2)
    $turned[$entry[0]]=[ordered]@{shapeId=[string]$entry[1].ID;cells=(Get-ShapeCells $entry[1] @('Width','Height','PinX','PinY','LocPinX','LocPinY','Angle'));transform=(Get-LineTransform $entry[1])}
   }
  }
  $document.SaveAs((Join-Path $directory 'quarter-turned.vsdx')) | Out-Null
  $page.Export((Join-Path $directory 'quarter-page.svg'))
  $evidence.Add('quarterTurn',$QuarterTurn)
  $evidence.Add('quarterTurned',$turned)
 }
 if($Flip -ne 'None'){
  foreach($shape in @($rectangle,$ellipse)){
   if($shape -and $FlipProtection -ne 'None'){
    if($FlipProtection -eq 'GuardAngle'){$shape.CellsU('Angle').FormulaU='GUARD('+ $shape.CellsU('Angle').FormulaU +')'}
    else{$shape.CellsU($FlipProtection).ResultIU=1.0}
   }
  }
  $document.SaveAs((Join-Path $directory 'flip-source.vsdx')) | Out-Null
  $flipped=[ordered]@{}
  $direction=if($Flip -eq 'Horizontal'){1}else{2}
  foreach($entry in @(@('rectangle',$rectangle),@('ellipse',$ellipse))){
   if($entry[1]){
    $selection=$page.CreateSelection(2,256,$entry[1])
    $selection.Flip($direction,2,$false)
    $flipped[$entry[0]]=[ordered]@{shapeId=[string]$entry[1].ID;cells=(Get-ShapeCells $entry[1] @('Width','Height','PinX','PinY','LocPinX','LocPinY','Angle','FlipX','FlipY'));transform=(Get-LineTransform $entry[1])}
   }
  }
  $document.SaveAs((Join-Path $directory 'flipped.vsdx')) | Out-Null
  $page.Export((Join-Path $directory 'flipped-page.svg'))
  $evidence.Add('flip',$Flip)
  $evidence.Add('flipped',$flipped)
 }
 if($DeleteAfterMove){
  foreach($shape in $shapes){$shape.Delete()}
  $retainedIds=@($control.ID)
  if($IncludeRectangle){$retainedIds+=,$rectangle.ID}
  if($IncludeEllipse){$retainedIds+=,$ellipse.ID}
  if($page.Shapes.Count -ne $retainedIds.Count){throw 'Deletion altered retained shapes.'}
  foreach($retained in $page.Shapes){if($retained.ID -notin $retainedIds){throw 'Deletion altered retained shapes.'}}
  $document.SaveAs((Join-Path $directory 'deleted.vsdx')) | Out-Null
  $evidence.Add('deletedShapeIds',@($cases | ForEach-Object shapeId))
  $evidence.Add('controlShapeId',[string]$control.ID)
 }
 [IO.File]::WriteAllText((Join-Path $directory 'evidence.json'),($evidence | ConvertTo-Json -Depth 7),[Text.UTF8Encoding]::new($false))
 Write-Output $directory
} finally {
 try {if($document){$document.Close()}} finally {$app.Quit()}
}
