param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$VerifyCore)
$ErrorActionPreference='Stop';New-Item -ItemType Directory -Force -Path $OutputDirectory|Out-Null;$dir=(Resolve-Path -LiteralPath $OutputDirectory).Path
$app=New-Object -ComObject Visio.InvisibleApp;$app.AlertResponse=7
try {
 if($app.Documents.Count -ne 0){throw 'Expected owned empty app'}
 $doc=$app.Documents.Add('');$cases=@()
 foreach($mode in @('numeric','date','string')) {
  $shape=$doc.Pages.Item(1).DrawRectangle(1,1+$cases.Count,5,1.7+$cases.Count)
  $shape.Text='Before X | Y After'
  $chars=$shape.Characters;$chars.Begin=7;$chars.End=8
  if($mode -eq 'numeric'){$chars.AddCustomFieldU('Width',2)}
  elseif($mode -eq 'date'){$chars.AddCustomFieldU('DATE(2026,1,1)',204)}
  else {$chars.AddCustomFieldU('PAGENAME()',37)}
  $at=$shape.Characters.Text.IndexOf('Y');$chars=$shape.Characters;$chars.Begin=$at;$chars.End=$at+1;$chars.AddCustomFieldU('Height',2)
  $cases += [ordered]@{id=$shape.ID;mode=$mode;text=$shape.Characters.Text;source=$shape.Text;valueFormula=$shape.CellsU('Fields.Value[1]').FormulaU;formatFormula=$shape.CellsU('Fields.Format[1]').FormulaU}
 }
 $doc.SaveAs((Join-Path $dir 'field-edit-native.vsdx'))|Out-Null
 [IO.File]::WriteAllText((Join-Path $dir 'field-edit-native.json'),([ordered]@{version=$app.Version;cases=$cases}|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
 $cases|ConvertTo-Json -Depth 6
 if($VerifyCore){
  $doc.Saved=$true;$doc.Close()
  & bun (Join-Path $PSScriptRoot 'visio-existing-field-write.ts') (Join-Path $dir 'field-edit-native.vsdx') $dir
  if($LASTEXITCODE -ne 0){throw 'Core field update generation failed'}
  $edits=([IO.File]::ReadAllText((Join-Path $dir 'core-edits.json'))|ConvertFrom-Json).cases
  foreach($edit in $edits){
   for($round=0;$round -lt 2;$round++){
    $doc=$app.Documents.Open([IO.Path]::GetFullPath($edit.path))
    try{
     $shape=$doc.Pages.Item(1).Shapes.ItemFromID($edit.id)
     if($shape.Characters.Text -cne $edit.text){throw "Native field update mismatch: $($edit.id): $($shape.Characters.Text) vs $($edit.text)"}
     if($shape.CellsU('Fields.Value[1]').FormulaU -cne $edit.formula){throw 'Value formula changed'}
     if($round -eq 0){$doc.Save()|Out-Null}
    }finally{$doc.Saved=$true;$doc.Close()}
   }
  }
  [IO.File]::WriteAllText((Join-Path $dir 'native-edits.json'),([ordered]@{version=$app.Version;result='native field updates matched through save/reopen';cases=$edits}|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
 }
} finally {foreach($doc in @($app.Documents)){$doc.Saved=$true;$doc.Close()};$app.Quit();[Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)|Out-Null}
