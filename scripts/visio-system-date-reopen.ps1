param(
 [Parameter(Mandatory=$true)][string]$CoreExport,
 [Parameter(Mandatory=$true)][string]$Evidence
)
$ErrorActionPreference='Stop'
$path=[IO.Path]::GetFullPath($CoreExport)
$cases=([IO.File]::ReadAllText([IO.Path]::GetFullPath($Evidence))|ConvertFrom-Json).cases
$app=New-Object -ComObject Visio.InvisibleApp;$app.AlertResponse=7
try {
 if($app.Documents.Count -ne 0){throw 'Expected owned empty application'}
 for($round=0;$round -lt 2;$round++) {
  $doc=$app.Documents.Open($path)
  try {
   foreach($case in $cases) {
    $shape=$doc.Pages.Item(1).Shapes.ItemFromID($case.id)
    if($shape.Characters.Text -cne $case.text){throw "Mismatch id $($case.id)"}
    if($shape.CellsU('Fields.Value').FormulaU -cne $case.valueFormula){throw 'Value formula changed'}
    if($shape.CellsU('Fields.Format').FormulaU -cne "FIELDPICTURE($($case.code))"){throw 'Picture changed'}
   }
   if($round -eq 0){$doc.Save()|Out-Null}
  } finally {$doc.Saved=$true;$doc.Close()}
 }
 'Native core-export open/save/reopen matched all text/formula/picture cases.'
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
