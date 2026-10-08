# Owned hidden Characters-range replacement references; no GUI or clipboard access.
param(
 [string]$OutputDirectory=(Join-Path $env:TEMP ('visio-text-replace-'+[guid]::NewGuid().ToString('N'))),
 [string]$CoreOutputPath
)
$ErrorActionPreference='Stop'
$directory=[IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $directory|Out-Null
$app=New-Object -ComObject Visio.InvisibleApp
$document=$null
try {
 if($CoreOutputPath) {
  $evidence=Get-Content -Raw -LiteralPath (Join-Path $directory 'evidence.json')|ConvertFrom-Json
  $document=$app.Documents.Open([IO.Path]::GetFullPath($CoreOutputPath))
  foreach($case in $evidence.cases) {
   $page=$document.Pages.ItemFromID([int]$case.pageId)
   $shape=$page.Shapes.ItemFromID([int]$case.shapeId)
   if($shape.Text -cne $case.nativeText){throw "Native reopen text mismatch on page $($case.pageId)"}
  }
  [ordered]@{version=$app.Version;accepted=$evidence.cases.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
 } else {
  $document=$app.Documents.Add('')
  $emoji=[char]::ConvertFromUtf32(0x1f600)
  $patches=@(
   @{text='aaa aa';query='aa';replacement='X';mode='all'},
   @{text=$emoji+'a '+$emoji+'a';query='a';replacement='$&';mode='all'},
   @{text="a`na`n";query='a';replacement="B`n";mode='all'},
   @{text='a a';query='a';replacement='';mode='all'},
   @{text='A a A';query='A';replacement='a';mode='current'},
   @{text='a a a';query='a';replacement='aaaa';mode='current'}
  )
  $cases=@()
  foreach($scale in @(1,2,0.5)){foreach($patch in $patches){
   $page=if($cases.Count -eq 0){$document.Pages.Item(1)}else{$document.Pages.Add()}
   $page.Name='Replace-'+[string]$cases.Count
   $page.PageSheet.CellsU('PageScale').FormulaU='1 in'
   $page.PageSheet.CellsU('DrawingScale').FormulaU=([string]$scale)+' in'
   $shape=$page.DrawRectangle(1,1,4,2);$shape.Text=$patch.text
   $cases+=[ordered]@{pageId=[string]$page.ID;shapeId=[string]$shape.ID;request=$patch}
  }}
  $document.SaveAs((Join-Path $directory 'original.vsdx'))|Out-Null
  foreach($case in $cases) {
   $shape=$document.Pages.ItemFromID([int]$case.pageId).Shapes.ItemFromID([int]$case.shapeId)
   $text=[string]$shape.Text;$patch=$case.request;$matches=@();$start=0
   while(($start=$text.IndexOf([string]$patch.query,$start,[StringComparison]::Ordinal)) -ge 0) {
    $matches+=@{pageId=$case.pageId;shapeId=$case.shapeId;start=$start;end=$start+$patch.query.Length}
    $start+=$patch.query.Length
   }
   $case.occurrences=$matches
   if($patch.mode -eq 'current'){$matches=@($matches[$matches.Count-1]);$case.current=$matches[0]}
   for($index=$matches.Count-1;$index -ge 0;$index--) {
    $characters=$shape.Characters;$characters.Begin=$matches[$index].start;$characters.End=$matches[$index].end
    $characters.Text=[string]$patch.replacement
   }
   $case.nativeText=[string]$shape.Text
  }
  $document.SaveAs((Join-Path $directory 'native.vsdx'))|Out-Null
  [ordered]@{version=$app.Version;cases=$cases}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
 }
} finally {if($document){$document.Saved=$true;$document.Close()};$app.Quit()}
Write-Output $directory
