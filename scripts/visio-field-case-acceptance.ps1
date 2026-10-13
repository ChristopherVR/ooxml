param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$candidates=Join-Path $OutputDirectory 'candidates.json'
& node (Join-Path $PSScriptRoot 'visio-field-case-candidates.mjs') $candidates
if ($LASTEXITCODE) {throw 'Candidate generation failed.'}
$inputData=Get-Content -LiteralPath $candidates -Raw|ConvertFrom-Json
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 $doc=$app.Documents.Add('')
 try {
  $cases=@()
  foreach($value in $inputData.inputs) {
   $entry=[ordered]@{input=$value}
   foreach($code in @(38,39)) {
    $shape=$doc.Pages.Item(1).DrawRectangle(0,0,2,1)
    $shape.Characters.AddCustomFieldU(('"'+$value+'"'),$code)
    $entry[([string]$code)]=$shape.Characters.Text
   }
   $cases += $entry
  }
  [ordered]@{version=$app.Version;fileVersion=(Get-Item -LiteralPath (Join-Path $app.Path 'VISIO.EXE')).VersionInfo.FileVersion;unicode=$inputData.unicode;cases=$cases}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $OutputDirectory 'field-case-native.json')
  "Recorded $($cases.Count) case chunks"
 } finally {$doc.Saved=$true;$doc.Close()}
 $fixture=$app.Documents.Add('')
 try {
  foreach($code in @(37,38,39)) {
   $master=$fixture.Masters.Add()
   $master.NameU='Field '+$code
   $copy=$master.Open()
   $shape=$copy.DrawRectangle(0,0,2,1)
   $shape.AddNamedRow(242,'Value',0)|Out-Null
   $shape.CellsU('User.Value').FormulaU='"Placeholder"'
   $shape.Characters.AddCustomFieldU('User.Value',$code)
   $copy.Close()
   $instance=$fixture.Pages.Item(1).Drop($master,3,$code-36)
   $instance.CellsU('User.Value').FormulaU='"Straße café İΣ"'
  }
  $fixture.SaveAs((Join-Path $OutputDirectory 'field-case-native.vsdx'))|Out-Null
 } finally {$fixture.Saved=$true;$fixture.Close()}
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
