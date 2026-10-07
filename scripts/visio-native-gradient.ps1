# Shared native-oracle setup. The caller owns the invisible application and document.
function Set-VisioNativeGradient {
 param($Shape, [ValidateSet('Fill','Line')][string]$Paint, [string]$Angle, [string]$Foreground, [string]$Background,
  [string]$ForegroundTransparency, [string]$BackgroundTransparency,
  [ValidateRange(0,13)][int]$Direction = 0)
 $Shape.CellsU($Paint+'GradientEnabled').FormulaU='1'
 $Shape.CellsU($Paint+'GradientAngle').FormulaU=$Angle
 $Shape.CellsU($Paint+'GradientDir').FormulaU=[string]$Direction
 $Shape.CellsU('RotateGradientWithShape').FormulaU='1'
 $Shape.CellsU('UseGroupGradient').FormulaU='0'
 $section=if($Paint -eq 'Line'){248}else{249}
 if($Shape.SectionExists($section,0) -eq 0){$Shape.AddSection($section) | Out-Null}
 while($Shape.RowCount($section) -lt 2){$Shape.AddRow($section,-1,0) | Out-Null}
 $Shape.CellsSRC($section,0,0).FormulaU=$Foreground
 $Shape.CellsSRC($section,1,0).FormulaU=$Background
 $Shape.CellsSRC($section,0,1).FormulaU=$ForegroundTransparency
 $Shape.CellsSRC($section,1,1).FormulaU=$BackgroundTransparency
 $Shape.CellsSRC($section,0,2).FormulaU='0%'
 $Shape.CellsSRC($section,1,2).FormulaU='100%'
}

function Set-VisioNativeFillGradient {
 param($Shape, [string]$Angle, [string]$Foreground, [string]$Background,
  [string]$ForegroundTransparency, [string]$BackgroundTransparency,
  [ValidateRange(0,13)][int]$Direction = 0)
 Set-VisioNativeGradient $Shape 'Fill' $Angle $Foreground $Background $ForegroundTransparency $BackgroundTransparency $Direction
}
