# Owned native rich character-range references; no GUI or clipboard interaction.
param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-text-ranges-'+[guid]::NewGuid().ToString('N'))),
 [string]$CoreOutputPath
)
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $directory|Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
function Read-Rich($shape) {
 $characters=@();$text=[string]$shape.Text
 for($index=0;$index -lt $text.Length;$index++) {
  $range=$shape.Characters;$range.Begin=$index;$range.End=$index+1
  $charRow=$range.CharPropsRow(0);$paraRow=$range.ParaPropsRow(0)
  $characters+=[ordered]@{text=$range.Text;style=$shape.CellsSRC(3,$charRow,2).ResultIU;size=$shape.CellsSRC(3,$charRow,7).ResultIU;color=$shape.CellsSRC(3,$charRow,1).ResultStrU(0);alignment=$shape.CellsSRC(4,$paraRow,6).ResultIU}
 }
 [ordered]@{text=$text;characters=$characters}
}
function Normalize-Color($value) {
 if($value -match '^RGB\((\d+),\s*(\d+),\s*(\d+)\)$') { return ('#{0:x2}{1:x2}{2:x2}' -f [int]$Matches[1],[int]$Matches[2],[int]$Matches[3]) }
 if($value -eq '0'){return '#000000'}
 return $value.ToLowerInvariant()
}
try {
 $app.AlertResponse=7
 $app.EventsEnabled=0
 if($CoreOutputPath) {
  # The test writes effective saved-native properties, since in-memory row caches can be stale.
  $evidence=Get-Content -Raw -LiteralPath (Join-Path $directory 'core-reference.json')|ConvertFrom-Json
  $document=$app.Documents.OpenEx([IO.Path]::GetFullPath($CoreOutputPath),202)
  $reopened=@()
  foreach($case in $evidence.cases) {
   $actual=Read-Rich $document.Pages.ItemFromID([int]$case.command.pageId).Shapes.ItemFromID([int]$case.command.shapeId)
   if($actual.text -cne $case.native.text -or $actual.characters.Count -ne $case.native.characters.Count){throw 'Native rich text reopen mismatch'}
   for($index=0;$index -lt $actual.characters.Count;$index++) {
    $actual.characters[$index].color=Normalize-Color $actual.characters[$index].color
    foreach($name in @('text','color')){if($actual.characters[$index][$name] -cne $case.native.characters[$index].$name){throw "Native rich text $name mismatch: page $($case.command.pageId) index $index actual $($actual.characters[$index][$name]) expected $($case.native.characters[$index].$name)"}}
    foreach($name in @('style','size','alignment')){if([Math]::Abs($actual.characters[$index][$name]-$case.native.characters[$index].$name) -gt 1e-10){throw "Native rich text $name mismatch: page $($case.command.pageId) index $index actual $($actual.characters[$index][$name]) expected $($case.native.characters[$index].$name)"}}
   }
   $reopened+=[ordered]@{pageId=$case.command.pageId;shapeId=$case.command.shapeId;actual=$actual}
  }
  [ordered]@{version=$app.Version;accepted=$evidence.cases.Count;reference='Effective parsed saved-native text and styles';cases=$reopened}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
 } else {
  $document=$app.Documents.Add('')
  $specs=@(
   @(@{start=1;end=2;text='X'}),@(@{start=1;end=3;text='XY'}),
   @(@{start=2;end=3;text='LONG'}),@(@{start=2;end=3;text=''}),
   @(@{start=6;end=8;text='$&'}),
   @(@{start=0;end=1;text='Q'},@{start=3;end=4;text='R'},@{start=8;end=9;text='S'}),
   @(@{start=1;end=2;text='X'},@{start=2;end=3;text='Y'}),@(@{start=1;end=2;text='b'})
  )
  $cases=@()
  foreach($scale in @(1,2,0.5)){foreach($ranges in $specs){
   $page=if($cases.Count -eq 0){$document.Pages.Item(1)}else{$document.Pages.Add()};$page.Name='Range-'+[string]$cases.Count
   $page.PageSheet.CellsU('PageScale').FormulaU='1 in';$page.PageSheet.CellsU('DrawingScale').FormulaU=([string]$scale)+' in'
   $shape=$page.DrawRectangle(1,1,5,4);$shape.Text="abCD`nEFgh"
   $shape.CellsU('Char.Style').FormulaU='0';$shape.CellsU('Char.Font').FormulaU='FONT("Arial")';$shape.CellsU('Char.Size').FormulaU='12 pt';$shape.CellsU('Para.HorzAlign').FormulaU='0'
   $range=$shape.Characters;$range.Begin=2;$range.End=4;$range.CharProps(2)=2;$range.CharProps(7)=20
   $shape.CellsSRC(3,$range.CharPropsRow(0),1).FormulaU='RGB(255,0,0)'
   $range=$shape.Characters;$range.Begin=5;$range.End=7;$range.CharProps(2)=4;$range.CharProps(7)=16;$range.ParaProps(6)=2
   $shape.CellsSRC(3,$range.CharPropsRow(0),1).FormulaU='RGB(0,0,255)'
   $cases+=[ordered]@{command=[ordered]@{type='replace-text-ranges';pageId=[string]$page.ID;shapeId=[string]$shape.ID;expectedText=[string]$shape.Text;ranges=@($ranges)}}
  }}
  $document.SaveAs((Join-Path $directory 'original.vsdx'))|Out-Null
  foreach($case in $cases) {
   $shape=$document.Pages.ItemFromID([int]$case.command.pageId).Shapes.ItemFromID([int]$case.command.shapeId)
   $ranges=@($case.command.ranges)
   for($index=$ranges.Count-1;$index -ge 0;$index--){$range=$shape.Characters;$range.Begin=$ranges[$index].start;$range.End=$ranges[$index].end;$range.Text=[string]$ranges[$index].text}
   $case.native=Read-Rich $shape
  }
  $document.SaveAs((Join-Path $directory 'native.vsdx'))|Out-Null
  [ordered]@{version=$app.Version;cases=$cases}|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
 }
}finally{if($document){$document.Saved=$true;$document.Close();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)};$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
Write-Output $directory
