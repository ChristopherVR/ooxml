param([Parameter(Mandatory=$true)][string]$InputDirectory)
$ErrorActionPreference='Stop'
$InputDirectory=(Resolve-Path -LiteralPath $InputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse=7
try {
 foreach($name in @('numeric-field-native','numeric-field-boundary-native','numeric-field-general-native','numeric-field-fraction-native')) {
  $evidence=Get-Content -Raw -Encoding utf8 -LiteralPath (Join-Path $InputDirectory ($name+'.json'))|ConvertFrom-Json
  $source=Join-Path $InputDirectory ($name+'.vsdx')
  $resaved=Join-Path $InputDirectory ($name+'-resaved.vsdx')
  foreach($path in @($source,$resaved)) {
   $doc=$app.Documents.OpenEx($path,0)
   try {
    foreach($case in $evidence.cases) {
     $actual=$doc.Pages.Item(1).Shapes.ItemFromID($case.id).Characters.Text
     if($actual -cne $case.text) {throw "Native field mismatch: $name ID $($case.id), '$actual' versus '$($case.text)'"}
    }
    if($path -eq $source) {$doc.SaveAs($resaved)|Out-Null}
   } finally {$doc.Saved=$true;$doc.Close()}
  }
  "$name : $($evidence.cases.Count) native fields match through native resave/reopen"
 }
} finally {$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
