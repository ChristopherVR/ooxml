param([string]$OutputDirectory=(Join-Path $env:TEMP ('ooxml-word-nested-frames-'+[guid]::NewGuid().ToString('N'))))
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null
$root=Split-Path $PSScriptRoot -Parent
$cases=@(@{name='nested-frame-short';source='nested-table-reference'},@{name='nested-frame-two-rows';source='nested-table-reference';twoRows=$true},@{name='nested-frame-centered';source='nested-table-reference';center=$true},@{name='nested-frame-split';source='nested-split-late'})
$app=$null;$owned=$false;$records=@()
try{
 $app=New-Object -ComObject Word.Application;$owned=$app.Documents.Count-eq0
 if(-not$owned){throw 'New Word instance contains documents'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases){
  $inputPath=Join-Path $OutputDirectory ($case.name+'-input.docx')
  Copy-Item -LiteralPath (Join-Path $root ('src/core/docx/__fixtures__/table-footnotes/'+$case.source+'.docx')) -Destination $inputPath
  $zip=[IO.Compression.ZipFile]::Open($inputPath,[IO.Compression.ZipArchiveMode]::Update)
  try{
   $entry=$zip.GetEntry('word/document.xml');$reader=[IO.StreamReader]::new($entry.Open(),[Text.Encoding]::UTF8)
   try{$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
   $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
   $inner=$xml.SelectSingleNode('//w:tbl/w:tr/w:tc/w:tbl',$ns);if(-not$inner){throw 'Missing inner table'}
   $props=$inner.SelectSingleNode('w:tblPr',$ns)
   $existing=$props.SelectSingleNode('w:tblBorders',$ns);if($existing){[void]$props.RemoveChild($existing)}
   $borders=$xml.CreateElement('w','tblBorders',$ns.LookupNamespace('w'))
   foreach($side in @('top','left','bottom','right','insideH','insideV')){
    $edge=$xml.CreateElement('w',$side,$ns.LookupNamespace('w'))
    foreach($attribute in @(@('val','single'),@('sz','8'),@('color','FF0000'))){$a=$xml.CreateAttribute('w',$attribute[0],$ns.LookupNamespace('w'));$a.Value=$attribute[1];[void]$edge.Attributes.Append($a)}
    [void]$borders.AppendChild($edge)
   }
   [void]$props.AppendChild($borders)
   if($case.twoRows){$row=$inner.SelectSingleNode('w:tr',$ns).CloneNode($true);foreach($reference in $row.SelectNodes('.//w:footnoteReference',$ns)){[void]$reference.ParentNode.RemoveChild($reference)};foreach($text in $row.SelectNodes('.//w:t',$ns)){$text.InnerText=$text.InnerText.Replace('T1R0','T1R1')};[void]$inner.AppendChild($row)}
   $rowIndex=0
   foreach($row in $inner.SelectNodes('w:tr',$ns)){
    $cellIndex=0
    foreach($cell in $row.SelectNodes('w:tc',$ns)){
     $tcProps=$cell.SelectSingleNode('w:tcPr',$ns);$old=$tcProps.SelectSingleNode('w:shd',$ns);if($old){[void]$tcProps.RemoveChild($old)}
     $shade=$xml.CreateElement('w','shd',$ns.LookupNamespace('w'));$fill=$xml.CreateAttribute('w','fill',$ns.LookupNamespace('w'));$fill.Value=if($rowIndex-eq0){if($cellIndex-eq0){'FFFF00'}else{'00FFFF'}}else{if($cellIndex-eq0){'00FF00'}else{'FF00FF'}};[void]$shade.Attributes.Append($fill);[void]$tcProps.AppendChild($shade);$cellIndex++
    };$rowIndex++
   }
   if($case.center){
    foreach($width in $inner.SelectNodes('w:tblPr/w:tblW|w:tr/w:tc/w:tcPr/w:tcW',$ns)){$width.SetAttribute('w',$ns.LookupNamespace('w'),$(if($width.LocalName-eq'tblW'){'2400'}else{'1200'}))|Out-Null}
    foreach($grid in $inner.SelectNodes('w:tblGrid/w:gridCol',$ns)){$grid.SetAttribute('w',$ns.LookupNamespace('w'),'1200')|Out-Null}
    $jc=$props.SelectSingleNode('w:jc',$ns);if(-not$jc){$jc=$xml.CreateElement('w','jc',$ns.LookupNamespace('w'));[void]$props.AppendChild($jc)};$jc.SetAttribute('val',$ns.LookupNamespace('w'),'center')|Out-Null
   }
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
 Write-Output $OutputDirectory
}finally{if($app-and$owned){$app.Quit(0)};if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}}
