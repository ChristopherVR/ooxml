param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory|Out-Null
$outDir=(Resolve-Path -LiteralPath $OutputDirectory).Path
New-Item -ItemType Directory -Force -Path $outDir|Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 $doc=$app.Documents.Add('')
 $master=$doc.Masters.Add();$master.NameU='ISO date';$copy=$master.Open();$shape=$copy.DrawRectangle(0,0,2,1)
 $shape.AddNamedRow(242,'DateValue',0)|Out-Null
 $shape.CellsU('User.DateValue').FormulaU='DATE(2026,1,1)'
 $shape.Characters.AddCustomFieldU('User.DateValue',20);$copy.Close()
 $cases=@()
 foreach($formula in @('DATE(2026,10,13)','DATE(2024,2,29)','DATE(1900,2,28)','DATE(1900,3,1)','DATE(1899,12,30)','DATE(1899,12,31)','DATE(1899,12,29)','DATE(2026,10,13)+TIME(14,15,16)','DATE(2026,10,13)+TIME(23,59,59.999)','DATE(2026,12,31)+TIME(23,59,59)','DATE(2000,2,29)','DATE(2100,3,1)')) {
  $instance=$doc.Pages.Item(1).Drop($master,3,1+$cases.Count)
  $instance.CellsU('User.DateValue').FormulaU=$formula
  $cases += [ordered]@{id=$instance.ID;formula=$formula;text=$instance.Characters.Text;valueFormula=$instance.CellsU('Fields.Value').FormulaU;value=$instance.CellsU('Fields.Value').ResultStrU('')}
 }
 $doc.SaveAs((Join-Path $outDir 'field-iso-date-native.vsdx'))|Out-Null
 [ordered]@{version=$app.Version;cases=$cases}|ConvertTo-Json -Depth 5|Set-Content -Encoding utf8 -LiteralPath (Join-Path $outDir 'field-iso-date-native.json')
 $cases|ConvertTo-Json -Depth 4
} finally {foreach($doc in @($app.Documents)){$doc.Saved=$true;$doc.Close()};$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}