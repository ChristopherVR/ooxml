param([string]$OutputDirectory=(Join-Path $env:TEMP 'ooxml-native-word-interacting-note-columns'))
$ErrorActionPreference='Stop'
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
python (Join-Path $PSScriptRoot 'generate-word-interacting-note-columns.py') $OutputDirectory
if($LASTEXITCODE-ne0){throw 'Synthetic document generation failed'}
$cases=ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $OutputDirectory 'input-cases.json'),[Text.Encoding]::UTF8))
$app=$null;$owned=$false;$records=@()
try{
 $app=New-Object -ComObject Word.Application;$owned=$app.Documents.Count-eq0
 if(-not$owned){throw 'New Word instance contains documents'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases){
  $doc=$null
  try{
   $doc=$app.Documents.Open((Join-Path $OutputDirectory ($case.name+'-input.docx')),$false,$false)
   $doc.Sections.Item(1).Range.FootnoteOptions.LayoutColumns=$case.noteColumns
   $doc.SaveAs2((Join-Path $OutputDirectory ($case.name+'.docx')),16)
   $doc.ExportAsFixedFormat((Join-Path $OutputDirectory ($case.name+'.pdf')),17)
   $records+=@{name=$case.name;wordVersion=$app.Version;wordBuild=$app.Build;pages=$doc.ComputeStatistics(2);source=$case.source;noteLines=$case.noteLines;noteColumns=$case.noteColumns}
   Write-Output ('Recorded '+$case.name)
  }finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'native-cases.json'),(ConvertTo-Json -InputObject @($records) -Depth 6),[Text.UTF8Encoding]::new($false))
}finally{if($app-and$owned){$app.Quit(0)};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
