param([string]$OutputDirectory=(Join-Path $env:TEMP 'ooxml-word-nested-multi-row'),[switch]$VerifyCore,[ValidateSet('two','three')][string]$Profile='two')
$ErrorActionPreference='Stop'
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
python (Join-Path $PSScriptRoot 'generate-word-nested-multi-row.py') $OutputDirectory $Profile
if($LASTEXITCODE-ne0){throw 'Probe generation failed'}
$cases=Get-Content -Raw -Encoding utf8 (Join-Path $OutputDirectory 'probe-cases.json')|ConvertFrom-Json
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
   $records+=@{name=$case.name;lineCounts=$case.lineCounts;wordVersion=$app.Version;wordBuild=$app.Build;pages=$doc.ComputeStatistics(2)}
   Write-Output ('Recorded '+$case.name)
  }finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'native-cases.json'),(ConvertTo-Json -InputObject @($records) -Depth 8),[Text.UTF8Encoding]::new($false))
 if($VerifyCore){
  bun (Join-Path $PSScriptRoot 'write-word-nested-multi-row.ts') $OutputDirectory
  if($LASTEXITCODE-ne0){throw 'Core export failed'}
  $exports=Get-Content -Raw -Encoding utf8 (Join-Path $OutputDirectory 'core-export-cases.json')|ConvertFrom-Json
  $acceptance=@()
  foreach($export in $exports){
   $path=Join-Path $OutputDirectory ($export.name+'.docx')
   foreach($pass in 1..2){
    $doc=$null
    try{
     $doc=$app.Documents.Open($path,$false,$false)
     $pages=$doc.ComputeStatistics(2)
     if($pages-ne$export.pages-or-not$doc.Content.Text.Contains($export.outerText)){throw ('Native core acceptance failed '+$export.name)}
     $acceptance+=@{name=$export.name;pass=$pass;pages=$pages;outerText=$export.outerText}
     if($pass-eq1){$path=Join-Path $OutputDirectory ($export.name+'-native-reopen.docx');$doc.SaveAs2($path,16)}
    }finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
   }
  }
  [IO.File]::WriteAllText((Join-Path $OutputDirectory 'core-reopen-native.json'),(ConvertTo-Json -InputObject @($acceptance) -Depth 5),[Text.UTF8Encoding]::new($false))
 }

 python (Join-Path $PSScriptRoot 'measure-word-nested-table-frames.py') $OutputDirectory
 if($LASTEXITCODE-ne0){throw 'PDF measurement failed'}
}finally{if($app-and$owned){$app.Quit(0)};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
