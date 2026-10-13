param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory=(Resolve-Path -LiteralPath $OutputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 $doc=$app.Documents.Add('')
 try {
  $doc.Pages.Item(1).DrawRectangle(1,1,5,2).Text='Created: '
  $doc.SaveAs((Join-Path $OutputDirectory 'source.vsdx')) | Out-Null
 } finally {$doc.Saved=$true;$doc.Close()}
 & node (Join-Path $PSScriptRoot 'visio-date-field-write.mjs') (Join-Path $OutputDirectory 'source.vsdx') $OutputDirectory
 if($LASTEXITCODE -ne 0){throw 'Core field generation failed'}
 $evidence=Get-Content (Join-Path $OutputDirectory 'cases.json') -Raw | ConvertFrom-Json
 foreach($case in $evidence.cases) {
  for($round=0;$round -lt 2;$round++) {
   $doc=$app.Documents.Open($case.path)
   try {
    $actual=$doc.Pages.Item(1).Shapes.Item(1).Characters.Text
    if($actual -cne $case.text){throw "Native date field mismatch: $($case.name): $actual vs $($case.text)"}
    if($round -eq 0){$doc.Save()}
   } finally {$doc.Saved=$true;$doc.Close()}
  }
 }
 [ordered]@{version=$app.Version;timezone=$evidence.timezone;cases=$evidence.cases;result='native save/reopen matched'} | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 -LiteralPath (Join-Path $OutputDirectory 'native-result.json')
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null}