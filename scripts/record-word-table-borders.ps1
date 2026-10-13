param([string]$OutputDirectory=(Join-Path $env:TEMP ('ooxml-native-word-table-borders-'+[guid]::NewGuid().ToString('N'))),[string[]]$CaseNames)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Path $OutputDirectory -Force|Out-Null
$root=Split-Path $PSScriptRoot -Parent
$cases=@(
 @{name='ordinary-border-short';source='table-footnotes/short-table';size='24'},
 @{name='ordinary-border-half';source='table-footnotes/short-table';size='4'},
 @{name='ordinary-border-margin';source='table-footnotes/short-table';size='24';margin='120'},
 @{name='ordinary-border-mixed';source='table-footnotes/short-table';mixed=$true},
 @{name='enclosing-border-auto';source='nested-table-frames/nested-frame-mixed-outer-border';size='24'},
 @{name='enclosing-border-exact';source='nested-table-frames/nested-frame-mixed-outer-border';size='24';exact=30},
 @{name='enclosing-border-exact27';source='nested-table-frames/nested-frame-mixed-outer-border';size='24';exact=27},
 @{name='enclosing-border-exact28';source='nested-table-frames/nested-frame-mixed-outer-border';size='24';exact=28},
 @{name='enclosing-border-exact29';source='nested-table-frames/nested-frame-mixed-outer-border';size='24';exact=29},
 @{name='enclosing-border-clipped';source='nested-table-frames/nested-frame-mixed-outer-border';size='24';exact=22},
 @{name='ordinary-border-split-early';source='table-footnotes/split-early-reference';size='24'},
 @{name='ordinary-border-split-late';source='table-footnotes/split-late-reference';size='24'},
 @{name='enclosing-border-split';source='nested-table-frames/nested-frame-mixed-split';size='24'},
 @{name='ordinary-border-header';source='table-footnotes/repeated-header';size='24'}
)
if($CaseNames){$cases=@($cases|Where-Object{$_.name-in$CaseNames})}
$app=$null;$owned=$false;$records=@()
try{
 $app=New-Object -ComObject Word.Application;$owned=$app.Documents.Count-eq0
 if(-not$owned){throw 'New Word instance contains documents'}
 $app.Visible=$false;$app.DisplayAlerts=0
 foreach($case in $cases){
  $inputPath=Join-Path $OutputDirectory ($case.name+'-input.docx')
  Copy-Item -LiteralPath (Join-Path $root ('src/core/docx/__fixtures__/'+$case.source+'.docx')) -Destination $inputPath
  $zip=[IO.Compression.ZipFile]::Open($inputPath,[IO.Compression.ZipArchiveMode]::Update)
  try{
   $entry=$zip.GetEntry('word/document.xml');$reader=[IO.StreamReader]::new($entry.Open(),[Text.Encoding]::UTF8)
   try{$xml=[xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
   $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
   $table=$xml.SelectSingleNode('//w:body/w:tbl',$ns);if(-not$table){throw 'Missing body table'}
   $props=$table.SelectSingleNode('w:tblPr',$ns);$old=$props.SelectSingleNode('w:tblBorders',$ns);if($old){[void]$props.RemoveChild($old)}
   $borders=$xml.CreateElement('w','tblBorders',$ns.LookupNamespace('w'))
   foreach($side in @('top','left','bottom','right','insideH','insideV')){
    $edge=$xml.CreateElement('w',$side,$ns.LookupNamespace('w'));$size=if($case.mixed){@{top='16';left='8';bottom='4';right='24';insideH='12';insideV='16'}[$side]}else{$case.size}
    foreach($attribute in @(@('val','single'),@('sz',$size),@('color','FF0000'))){$a=$xml.CreateAttribute('w',$attribute[0],$ns.LookupNamespace('w'));$a.Value=$attribute[1];[void]$edge.Attributes.Append($a)}
    [void]$borders.AppendChild($edge)
   };[void]$props.AppendChild($borders)
   foreach($inner in $table.SelectNodes('.//w:tbl/w:tblPr/w:tblBorders/*',$ns)){$inner.SetAttribute('color',$ns.LookupNamespace('w'),'0000FF')|Out-Null}
   $rowIndex=0
   foreach($row in $table.SelectNodes('w:tr',$ns)){
    if($case.exact){$trProps=$row.SelectSingleNode('w:trPr',$ns);if(-not$trProps){$trProps=$xml.CreateElement('w','trPr',$ns.LookupNamespace('w'));[void]$row.PrependChild($trProps)};$old=$trProps.SelectSingleNode('w:trHeight',$ns);if($old){[void]$trProps.RemoveChild($old)};$height=$xml.CreateElement('w','trHeight',$ns.LookupNamespace('w'));$height.SetAttribute('val',$ns.LookupNamespace('w'),[string]($case.exact*20))|Out-Null;$height.SetAttribute('hRule',$ns.LookupNamespace('w'),'exact')|Out-Null;[void]$trProps.AppendChild($height)}
    $cellIndex=0
    foreach($cell in $row.SelectNodes('w:tc',$ns)){
     $tcProps=$cell.SelectSingleNode('w:tcPr',$ns);$old=$tcProps.SelectSingleNode('w:shd',$ns);if($old){[void]$tcProps.RemoveChild($old)}
     $shade=$xml.CreateElement('w','shd',$ns.LookupNamespace('w'));$fill=$xml.CreateAttribute('w','fill',$ns.LookupNamespace('w'));$fill.Value=if($rowIndex-eq0){if($cellIndex-eq0){'FFFF00'}else{'00FFFF'}}else{if($cellIndex-eq0){'00FF00'}else{'FF00FF'}};[void]$shade.Attributes.Append($fill);[void]$tcProps.AppendChild($shade)
     if($case.margin){$old=$tcProps.SelectSingleNode('w:tcMar',$ns);if($old){[void]$tcProps.RemoveChild($old)};$margins=$xml.CreateElement('w','tcMar',$ns.LookupNamespace('w'));foreach($side in @('top','left','bottom','right')){$m=$xml.CreateElement('w',$side,$ns.LookupNamespace('w'));$m.SetAttribute('w',$ns.LookupNamespace('w'),$case.margin)|Out-Null;$m.SetAttribute('type',$ns.LookupNamespace('w'),'dxa')|Out-Null;[void]$margins.AppendChild($m)};[void]$tcProps.AppendChild($margins)}
     $cellIndex++
    };$rowIndex++
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
