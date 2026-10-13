param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 $doc=$app.Documents.Add('')
 try {
  $master=$doc.Masters.Add()
  $master.NameU='Synthetic field'
  $copy=$master.Open()
  $shape=$copy.DrawRectangle(0,0,2,1)
  $shape.AddNamedRow(242,'ValueText',0)|Out-Null
  $shape.CellsU('User.ValueText').FormulaU='"Abc"'
  $shape.Characters.AddCustomFieldU('User.ValueText',0)
  $copy.Close()
  $out=@()
  foreach($value in @('42','Ready','73.50','')) {
   $instance=$doc.Pages.Item(1).Drop($master,3,1+$out.Count)
   $instance.CellsU('User.ValueText').FormulaU='"'+$value+'"'
   $out += [ordered]@{id=$instance.ID;value=$value;text=$instance.Characters.Text;format=$instance.CellsU('Fields.Format').FormulaU}
  }
  $doc.SaveAs((Join-Path $OutputDirectory 'inherited-field-native.vsdx'))
  [ordered]@{version=$app.Version;cases=$out}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputDirectory 'evidence.json')
  $out|ConvertTo-Json
 } finally {$doc.Saved=$true;$doc.Close()}
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}