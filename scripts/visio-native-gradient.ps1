# Shared native-oracle setup. The caller owns the invisible application and document.
function Set-VisioNativeFillGradient {
 param($Shape, [string]$Angle, [string]$Foreground, [string]$Background,
  [string]$ForegroundTransparency, [string]$BackgroundTransparency,
  [ValidateRange(0,13)][int]$Direction = 0)
 $Shape.CellsU('FillGradientEnabled').FormulaU='1'
 $Shape.CellsU('FillGradientAngle').FormulaU=$Angle
 $Shape.CellsU('FillGradientDir').FormulaU=[string]$Direction
 $Shape.CellsU('RotateGradientWithShape').FormulaU='1'
 $Shape.CellsU('UseGroupGradient').FormulaU='0'
 if($Shape.SectionExists(249,0) -eq 0){$Shape.AddSection(249) | Out-Null}
 while($Shape.RowCount(249) -lt 2){$Shape.AddRow(249,-1,0) | Out-Null}
 $Shape.CellsSRC(249,0,0).FormulaU=$Foreground
 $Shape.CellsSRC(249,1,0).FormulaU=$Background
 $Shape.CellsSRC(249,0,1).FormulaU=$ForegroundTransparency
 $Shape.CellsSRC(249,1,1).FormulaU=$BackgroundTransparency
 $Shape.CellsSRC(249,0,2).FormulaU='0%'
 $Shape.CellsSRC(249,1,2).FormulaU='100%'
}
