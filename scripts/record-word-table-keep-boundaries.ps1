param([string]$OutputDirectory=(Join-Path $env:TEMP 'ooxml-word-keep-boundary-native'))
$ErrorActionPreference='Stop'
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
python (Join-Path $PSScriptRoot 'generate-word-table-keep-boundaries.py') $OutputDirectory
if($LASTEXITCODE-ne0){throw 'Boundary generation failed'}
$cases=[IO.File]::ReadAllText((Join-Path $OutputDirectory 'probe-cases.json'))|ConvertFrom-Json
$app=$null;$owned=$false;$records=@()
try{
 $app=New-Object -ComObject Word.Application;$owned=$app.Documents.Count-eq0
 if(-not$owned){throw 'New Word instance contains documents'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases){
  $doc=$null
  try{
   $doc=$app.Documents.Open((Join-Path $OutputDirectory ($case.name+'-input.docx')),$false,$false)
   $doc.SaveAs2((Join-Path $OutputDirectory ($case.name+'.docx')),16)
   $doc.ExportAsFixedFormat((Join-Path $OutputDirectory ($case.name+'.pdf')),17)
   $records+=@{name=$case.name;wordVersion=$app.Version;wordBuild=$app.Build;pages=$doc.ComputeStatistics(2)}
  }finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'native-cases.json'),(ConvertTo-Json -InputObject @($records) -Depth 8),[Text.UTF8Encoding]::new($false))
}finally{
 if($app){if($owned){$app.Quit()};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
python (Join-Path $PSScriptRoot 'measure-word-nested-table-frames.py') $OutputDirectory
if($LASTEXITCODE-ne0){throw 'Boundary measurement failed'}
python (Join-Path $PSScriptRoot 'measure-word-table-keep-boundaries.py') $OutputDirectory
if($LASTEXITCODE-ne0){throw 'Compact boundary measurement failed'}
