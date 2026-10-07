# Shared native COM measurement helpers for Visio capture scripts.
function Get-ShapeCells($shape,[string[]]$names){
 $result=[ordered]@{}
 foreach($name in $names){$cell=$shape.CellsU($name);$result[$name]=[ordered]@{formula=$cell.FormulaU;value=$cell.ResultIU}}
 return $result
}
function Get-LineTransform($shape){
 $x0=0.0;$y0=0.0;$xx=0.0;$yx=0.0;$xy=0.0;$yy=0.0
 $shape.XYToPage(0,0,[ref]$x0,[ref]$y0)
 $shape.XYToPage(1,0,[ref]$xx,[ref]$yx)
 $shape.XYToPage(0,1,[ref]$xy,[ref]$yy)
 return @(($xx-$x0),($yx-$y0),($xy-$x0),($yy-$y0),$x0,$y0)
}
