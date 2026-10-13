param([string]$OutputDirectory=(Join-Path $env:TEMP 'ooxml-native-word-independent-inner-notes'))
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null
$OutputDirectory=[IO.Path]::GetFullPath($OutputDirectory)
$root=Split-Path $PSScriptRoot -Parent
$cases=@(@{name='inner-note-early';cell=0;line=0},@{name='inner-note-second-late';cell=1;line=15},@{name='inner-note-long-neighbor';cell=0;line=15;long=1},@{name='inner-note-long-reference';cell=0;line=25;long=0})
$app=$null;$owned=$false;$records=@()
try{
 $app=New-Object -ComObject Word.Application;$owned=$app.Documents.Count-eq0
 if(-not$owned){throw 'New Word instance contains documents'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases){
  $inputPath=Join-Path $OutputDirectory ($case.name+'-input.docx')
  Copy-Item -LiteralPath (Join-Path $root 'src/core/docx/__fixtures__/ordinary-table-borders/enclosing-border-split.docx') -Destination $inputPath
  $zip=[IO.Compression.ZipFile]::Open($inputPath,[IO.Compression.ZipArchiveMode]::Update)
  try{
   $entry=$zip.GetEntry('word/document.xml');$reader=[IO.StreamReader]::new($entry.Open(),[Text.Encoding]::UTF8)
   try{$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
   $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
   $inner=$xml.SelectSingleNode('//w:body/w:tbl/w:tr/w:tc/w:tbl',$ns)
   $paragraphs=@($inner.SelectNodes('w:tr/w:tc/w:p',$ns))
   $ref=$inner.SelectSingleNode('.//w:r[w:footnoteReference]',$ns);[void]$ref.ParentNode.RemoveChild($ref)
   if($case.ContainsKey('long')){
    $paragraph=$paragraphs[$case.long];$last=$paragraph.SelectSingleNode('w:r[w:t[contains(.,"L19")]]',$ns)
    foreach($line in 20..29){$run=$last.CloneNode($true);$run.SelectSingleNode('w:t',$ns).InnerText=('T1R0C'+$case.long+'L'+$line.ToString('00'));[void]$paragraph.AppendChild($run)}
   }
   $target=$paragraphs[$case.cell].SelectSingleNode('w:r[w:t[contains(.,"L'+$case.line.ToString('00')+'")]]',$ns)
   if(-not$target){throw 'Missing marker line'}
   [void]$target.ParentNode.InsertAfter($ref,$target)
   $entry.Delete();$entry=$zip.CreateEntry('word/document.xml');$writer=[IO.StreamWriter]::new($entry.Open(),[Text.UTF8Encoding]::new($false));try{$writer.Write($xml.OuterXml)}finally{$writer.Dispose()}
  }finally{$zip.Dispose()}
  $doc=$null
  try{
   $doc=$app.Documents.Open($inputPath,$false,$false)
   $doc.SaveAs2((Join-Path $OutputDirectory ($case.name+'.docx')),16)
   $doc.ExportAsFixedFormat((Join-Path $OutputDirectory ($case.name+'.pdf')),17)
   $records+=@{name=$case.name;wordVersion=$app.Version;wordBuild=$app.Build;pages=$doc.ComputeStatistics(2)}
   Write-Output ('Recorded '+$case.name)
  }finally{if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}}
 }
 [IO.File]::WriteAllText((Join-Path $OutputDirectory 'native-cases.json'),(ConvertTo-Json -InputObject @($records) -Depth 5),[Text.UTF8Encoding]::new($false))
}finally{if($app-and$owned){$app.Quit(0)};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
