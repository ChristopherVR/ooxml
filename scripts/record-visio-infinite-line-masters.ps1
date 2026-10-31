# Synthetic masters only. Row type 141 is visTagInfiniteLine:
# https://learn.microsoft.com/en-us/office/vba/api/visio.visrowtags
param([Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$parityOut=(Resolve-Path -LiteralPath $OutputDirectory).Path
$parityApp=New-Object -ComObject Visio.InvisibleApp
$parityApp.AlertResponse=7
$parityDoc=$null
try {
 $parityDoc=$parityApp.Documents.Add('')
 $parityPage=$parityDoc.Pages.Item(1)
 $parityPage.PageSheet.CellsU('PageWidth').FormulaU='4 in'
 $parityPage.PageSheet.CellsU('PageHeight').FormulaU='3 in'
 foreach($parityName in @('Diagonal','Horizontal','Vertical')) {
  $parityMaster=$parityDoc.Masters.Add()
  $parityMaster.NameU=$parityName
  $parityCopy=$parityMaster.Open()
  try {
   $parityCopy.PageSheet.CellsU('PageWidth').FormulaU='4 in'
   $parityCopy.PageSheet.CellsU('PageHeight').FormulaU='3 in'
   $parityShape=$parityCopy.DrawRectangle(0,0,1,1)
   $parityShape.DeleteSection(10)
   $parityShape.AddSection(10)|Out-Null
   $parityShape.CellsU('Geometry1.NoFill').FormulaU='1'
   $parityRow=$parityShape.AddRow(10,1,141)
   $parityPoints=switch($parityName){
    'Diagonal' {@(0,0,1,1)}
    'Horizontal' {@(0,0.5,1,0.5)}
    'Vertical' {@(0.5,0,0.5,1)}
   }
   for($parityIndex=0;$parityIndex -lt 4;$parityIndex++) {
    $parityShape.CellsSRC(10,$parityRow,$parityIndex).FormulaU=$parityPoints[$parityIndex].ToString([cultureinfo]::InvariantCulture)+' in'
   }
  } finally {$parityCopy.Close()}
  $parityInstance=$parityPage.Drop($parityMaster,0.5,0.5)
  $parityPage.Export((Join-Path $parityOut "$parityName.svg"))
  $parityInstance.Delete()
 }
 $parityDoc.SaveAs((Join-Path $parityOut 'infinite-line-masters-native.vsdx'))
 Write-Output "Visio $($parityApp.Version): $($parityDoc.Masters.Count) synthetic infinite-line masters"
} finally {
 if($parityDoc){$parityDoc.Saved=$true;$parityDoc.Close()}
 $parityApp.Quit()
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($parityApp)|Out-Null
}

# Remove local user names from the generated fixture's package properties.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$parityZip=[IO.Compression.ZipFile]::Open((Join-Path $parityOut 'infinite-line-masters-native.vsdx'),'Update')
try {
 foreach($parityName in @('docProps/core.xml','docProps/app.xml')) {
  $parityEntry=$parityZip.GetEntry($parityName)
  if(-not $parityEntry){continue}
  $parityReader=[IO.StreamReader]::new($parityEntry.Open())
  try {$parityXml=[xml]$parityReader.ReadToEnd()} finally {$parityReader.Dispose()}
  foreach($parityNode in $parityXml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) {$parityNode.InnerText='Parity Author'}
  foreach($parityNode in $parityXml.SelectNodes('//*[local-name()="Company" or local-name()="Manager"]')) {$parityNode.InnerText=''}
  $parityEntry.Delete()
  $parityWriter=[IO.StreamWriter]::new($parityZip.CreateEntry($parityName).Open(),[Text.UTF8Encoding]::new($false))
  try {$parityWriter.Write($parityXml.OuterXml)} finally {$parityWriter.Dispose()}
 }
} finally {$parityZip.Dispose()}
