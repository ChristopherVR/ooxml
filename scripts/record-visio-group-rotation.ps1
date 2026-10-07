param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-group-rotation-'+[guid]::NewGuid().ToString('N'))),
 [double]$DrawingScale=1,
 [double]$PageScale=1,
 [double]$RotationDegrees=30,
 [switch]$Nested,
 [switch]$OffCentrePin,
 [ValidateSet('None','Left','Right')][string]$QuarterTurn='None'
)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'visio-capture-geometry.ps1')
$directory=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $directory){throw 'Use a fresh output directory.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
function Get-CaptureTree($shape){
 $names=@('PinX','PinY','Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY')
 $children=@()
 for($i=1;$i -le $shape.Shapes.Count;$i++){$children+=,(Get-CaptureTree $shape.Shapes.Item($i))}
 return [ordered]@{id=[string]$shape.ID;cells=(Get-ShapeCells $shape $names);transform=(Get-LineTransform $shape);children=$children}
}
try {
 $document=$app.Documents.Add('')
 $page=$document.Pages.Item(1)
 $page.PageSheet.CellsU('DrawingScale').ResultIU=$DrawingScale
 $page.PageSheet.CellsU('PageScale').ResultIU=$PageScale
 $coordinates=$DrawingScale/$PageScale
 $rectangle=$page.DrawRectangle(1*$coordinates,1*$coordinates,3*$coordinates,2*$coordinates)
 $ellipse=$page.DrawOval(4*$coordinates,2*$coordinates,6*$coordinates,4*$coordinates)
 $selection=$page.CreateSelection(0,256)
 $selection.Select($rectangle,2)
 $selection.Select($ellipse,2)
 $group=$selection.Group()
 if($Nested){
  $third=$page.DrawRectangle(2*$coordinates,5*$coordinates,4*$coordinates,6*$coordinates)
  $selection=$page.CreateSelection(0,256)
  $selection.Select($group,2)
  $selection.Select($third,2)
  $group=$selection.Group()
 }
 if($OffCentrePin){
  $group.CellsU('LocPinX').ResultIU=$group.CellsU('Width').ResultIU*0.25
  $group.CellsU('LocPinY').ResultIU=$group.CellsU('Height').ResultIU*0.75
 }
 $source=Get-CaptureTree $group
 $document.SaveAs((Join-Path $directory 'source.vsdx')) | Out-Null
 $page.Export((Join-Path $directory 'source-page.svg'))
 if($QuarterTurn -eq 'None'){$group.CellsU('Angle').ResultIU=$RotationDegrees*[Math]::PI/180}
 else {
  $selection=$page.CreateSelection(0,256)
  $selection.Select($group,2)
  $degrees=if($QuarterTurn -eq 'Left'){90}else{-90}
  $selection.Rotate($degrees,'deg',$false,2)
 }
 $rotated=Get-CaptureTree $group
 $document.SaveAs((Join-Path $directory 'rotated.vsdx')) | Out-Null
 $page.Export((Join-Path $directory 'rotated-page.svg'))
 $evidence=[ordered]@{application='Microsoft Visio';version=$app.Version;drawingScale=$DrawingScale;pageScale=$PageScale;quarterTurn=$QuarterTurn;source=$source;rotated=$rotated}
 [IO.File]::WriteAllText((Join-Path $directory 'evidence.json'),($evidence | ConvertTo-Json -Depth 16),[Text.UTF8Encoding]::new($false))
 Write-Output $directory
} finally {
 try {if($document){$document.Close()}} finally {$app.Quit()}
}
