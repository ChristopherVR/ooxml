param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$BoundaryCases)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$values = if ($BoundaryCases) { @('1.005','-1.005','2.675','-0.0001','0','-0.1','0.001','0.00001','1000000000000000') } else { @('73.456','-2.3456','73.456 in','73.456 mm','7.3456 cm','2.3456 ft','45.6789 deg','0.73456 rad') }
$prefix = if ($BoundaryCases) { 'numeric-field-boundary-native' } else { 'numeric-field-native' }
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 $doc=$app.Documents.Add('')
 try {
  $out=@()
  foreach($code in 2..9) {
   foreach($formula in $values) {
    $master=$doc.Masters.Add()
    $master.NameU='Numeric field '+$out.Count
    $copy=$master.Open()
    $shape=$copy.DrawRectangle(0,0,2,1)
    $shape.AddNamedRow(242,'ValueNumber',0)|Out-Null
    $unit=$formula -replace '^[+-]?[\d.]+\s*',''
    $shape.CellsU('User.ValueNumber').FormulaU='1 '+$unit
    $shape.Characters.AddCustomFieldU('User.ValueNumber',$code)
    $copy.Close()
    $instance=$doc.Pages.Item(1).Drop($master,3,1+$out.Count)
    $instance.CellsU('User.ValueNumber').FormulaU=$formula
    $out += [ordered]@{id=$instance.ID;code=$code;formula=$formula;text=$instance.Characters.Text;format=$instance.CellsU('Fields.Format').FormulaU}
   }
  }
  $doc.SaveAs((Join-Path $OutputDirectory ($prefix+'.vsdx')))
  [ordered]@{version=$app.Version;cases=$out}|ConvertTo-Json -Depth 5|Set-Content -Encoding utf8 -LiteralPath (Join-Path $OutputDirectory ($prefix+'.json'))
 } finally {$doc.Saved=$true;$doc.Close()}
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}