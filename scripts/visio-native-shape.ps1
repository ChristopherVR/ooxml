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

function New-VisioNativeFillShape {
 param($Page, [ValidateSet('rectangle','ellipse','triangle','notched')][string]$Kind)
 if($Kind -eq 'ellipse'){return $Page.DrawOval(1,1,3,2)}
 if($Kind -eq 'triangle' -or $Kind -eq 'notched'){
  $points=[double[]]$(if($Kind -eq 'triangle'){@(1,1,3,1,2,2,1,1)}else{@(1,1,3,1,3,1.4,2,1.4,2,2,1,2,1,1)})
  return $Page.DrawPolyline([ref]$points,0)
 }
 $shape=$Page.DrawRectangle(1,1,3,2)
 return $shape
}
