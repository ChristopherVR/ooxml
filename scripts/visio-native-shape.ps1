# Explicit cached rows avoid compressed PolylineTo data in geometry probes.
function Set-VisioNativePolylineGeometry {
 param($Shape, [double[]]$Points, [bool]$NoFill = $false)
 if($Points.Count -lt 4 -or $Points.Count % 2){throw 'Expected x/y vertex pairs.'}
 $Shape.DeleteSection(10)
 $Shape.AddSection(10) | Out-Null
 $Shape.CellsU('Geometry1.NoFill').FormulaU=if($NoFill){'1'}else{'0'}
 for($vertex=0;$vertex -lt $Points.Count/2;$vertex++){
  $tag=if($vertex -eq 0){138}else{139}
  $row=$Shape.AddRow(10,$vertex+1,$tag)
  $Shape.CellsSRC(10,$row,0).FormulaU=$Points[2*$vertex].ToString([cultureinfo]::InvariantCulture)+' in'
  $Shape.CellsSRC(10,$row,1).FormulaU=$Points[2*$vertex+1].ToString([cultureinfo]::InvariantCulture)+' in'
 }
}

function Get-VisioNativeFillPoints {
 param([string]$Kind)
 switch($Kind){
  'triangle' {return [double[]]@(1,1,3,1,2,2,1,1)}
  'notched' {return [double[]]@(1,1,3,1,3,1.4,2,1.4,2,2,1,2,1,1)}
  'pentagon' {return [double[]]@(1,1,3,1,3,1.6,2,2,1,1.6,1,1)}
  'chevron' {return [double[]]@(1,1,2,1,3,1.5,2,2,1,2,2,1.5,1,1)}
  'ushape' {return [double[]]@(1,1,3,1,3,2,2.5,2,2.5,1.25,1.5,1.25,1.5,2,1,2,1,1)}
  'star' {return [double[]]@(2,2,2.2,1.6,3,1.65,2.3,1.3,2.6,1,2,1.2,1.4,1,1.7,1.3,1,1.65,1.8,1.6,2,2)}
 }
 return [double[]]@(1,1,3,1,3,2,1,2,1,1)
}

function New-VisioNativeFillShape {
 param($Page, [ValidateSet('rectangle','ellipse','triangle','notched','pentagon','chevron','ushape','star')][string]$Kind)
 if($Kind -eq 'ellipse'){return $Page.DrawOval(1,1,3,2)}
 if($Kind -ne 'rectangle'){
  $points=[double[]](Get-VisioNativeFillPoints $Kind)
  return $Page.DrawPolyline([ref]$points,0)
 }
 $shape=$Page.DrawRectangle(1,1,3,2)
 return $shape
}
