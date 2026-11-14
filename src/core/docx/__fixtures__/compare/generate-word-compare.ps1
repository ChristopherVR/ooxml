# Builds synthetic original/revised pairs with Word 16 and saves Word's own Review > Compare
# result for each (Application.CompareDocuments, new destination), then anonymizes package
# properties. The results are the expectations compare.word.test.ts checks compareDocuments against.
$ErrorActionPreference='Stop'
$cmpWord=New-Object -ComObject Word.Application
$cmpWord.Visible=$false
$cmpWord.DisplayAlerts=0
$cmpAuthor='Compare Author'

function New-CmpDoc([string]$text) {
 $doc=$cmpWord.Documents.Add()
 $doc.Content.Text=$text
 return $doc
}
function Save-CmpDoc($doc,[string]$name) {
 Write-Output "saving $name"
 # A path left as a PowerShell object (not [string]) makes SaveAs2 wait on a hidden dialog.
 $doc.SaveAs2([string](Join-Path $PSScriptRoot $name),16)
 $doc.Close(0)
}
function Compare-CmpPair([string]$case,[int]$granularity) {
 $original=$cmpWord.Documents.Open([string](Join-Path $PSScriptRoot "$case-original.docx"))
 $revised=$cmpWord.Documents.Open([string](Join-Path $PSScriptRoot "$case-revised.docx"))
 # Destination 2 = new document; formatting, case, whitespace, tables, headers, footnotes,
 # text boxes, fields and comments compared; moves off; warnings ignored.
 $result=$cmpWord.CompareDocuments($original,$revised,2,$granularity,$true,$true,$true,$true,$true,$true,$true,$true,$true,$false,$cmpAuthor,$true)
 Write-Output "$case : $($result.Revisions.Count) revisions"
 $result.SaveAs2([string](Join-Path $PSScriptRoot "$case-word.docx"),16)
 $result.Close(0)
 $original.Close(0)
 $revised.Close(0)
}

try {
 # 1. Word-level edits inside paragraphs, a paragraph inserted and one deleted.
 Save-CmpDoc (New-CmpDoc "The quick brown fox jumps over the lazy dog.`rThis paragraph will be removed.`rClosing words stay here.") 'words-original.docx'
 Save-CmpDoc (New-CmpDoc "The slow brown fox leaps over the lazy cat.`rClosing words stay here.`rA brand new final paragraph.") 'words-revised.docx'
 Compare-CmpPair 'words' 1

 # 2. A formatting change (one word bolded, one italic removed) with a text change.
 $doc=New-CmpDoc "Plain words become bold here.`rItalic text loses its slant."
 $doc.Paragraphs.Item(2).Range.Words.Item(1).Italic=-1
 Save-CmpDoc $doc 'format-original.docx'
 $doc=New-CmpDoc "Plain words become bold here.`rItalic text loses its slant today."
 $doc.Paragraphs.Item(1).Range.Words.Item(4).Bold=-1
 Save-CmpDoc $doc 'format-revised.docx'
 Compare-CmpPair 'format' 1

 # 3. Both inputs carry tracked changes of their own: Word compares them as accepted.
 $doc=New-CmpDoc "Alpha beta gamma.`rDelta epsilon."
 $doc.TrackRevisions=$true
 $doc.Paragraphs.Item(1).Range.Words.Item(2).Text='BETA '
 $doc.TrackRevisions=$false
 Save-CmpDoc $doc 'tracked-original.docx'
 $doc=New-CmpDoc "Alpha beta gamma.`rDelta epsilon."
 $doc.TrackRevisions=$true
 $doc.Paragraphs.Item(2).Range.InsertAfter('Zeta.')
 $doc.Paragraphs.Item(1).Range.Words.Item(3).Text='GAMMA'
 $doc.TrackRevisions=$false
 Save-CmpDoc $doc 'tracked-revised.docx'
 Compare-CmpPair 'tracked' 1

 # 4. Character granularity: only the changed letters are marked.
 Save-CmpDoc (New-CmpDoc "The colour of the harbour is grey.") 'chars-original.docx'
 Save-CmpDoc (New-CmpDoc "The color of the harbor is gray.") 'chars-revised.docx'
 Compare-CmpPair 'chars' 0

 # 5. A table: one cell edited, a row inserted, a row deleted.
 foreach($side in @('original','revised')) {
  $doc=New-CmpDoc "Before the table."
  $doc.Content.InsertParagraphAfter()
  $range=$doc.Paragraphs.Item(2).Range
  $cells=if($side -eq 'original'){@('Name','Score','Ann','10','Bob','20','Eve','50')}else{@('Name','Score','Ann','15','Eve','50','Cy','30')}
  $table=$doc.Tables.Add($range,4,2)
  for($i=0;$i -lt 8;$i++){$table.Cell([math]::Floor($i/2)+1,($i%2)+1).Range.Text=$cells[$i]}
  Save-CmpDoc $doc "table-$side.docx"
 }
 Compare-CmpPair 'table' 1
} finally {
 $cmpWord.Quit()
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($cmpWord)|Out-Null
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
foreach($cmpFile in Get-ChildItem -Path $PSScriptRoot -Filter '*.docx') {
 $cmpArchive=[IO.Compression.ZipFile]::Open($cmpFile.FullName,'Update')
 try {
  # Tracked edits in the inputs carry Word's user name: replace every revision author.
  foreach($cmpEntry in @($cmpArchive.Entries | Where-Object {$_.FullName -like 'word/*.xml'})) {
   $cmpReader=[IO.StreamReader]::new($cmpEntry.Open())
   try {$cmpText=$cmpReader.ReadToEnd()} finally {$cmpReader.Dispose()}
   $cmpNew=[regex]::Replace($cmpText,'w:author="[^"]*"',"w:author=`"$cmpAuthor`"")
   $cmpNew=[regex]::Replace($cmpNew,'w15:author="[^"]*"',"w15:author=`"$cmpAuthor`"")
   $cmpNew=[regex]::Replace($cmpNew,'w15:userId="[^"]*"',"w15:userId=`"$cmpAuthor`"")
   if($cmpNew -eq $cmpText){continue}
   $cmpEntryName=$cmpEntry.FullName
   $cmpEntry.Delete()
   $cmpWriter=[IO.StreamWriter]::new($cmpArchive.CreateEntry($cmpEntryName).Open(),[Text.UTF8Encoding]::new($false))
   try {$cmpWriter.Write($cmpNew)} finally {$cmpWriter.Dispose()}
  }
  foreach($cmpName in @('docProps/core.xml','docProps/app.xml')) {
   $cmpEntry=$cmpArchive.GetEntry($cmpName)
   if(-not $cmpEntry){continue}
   $cmpReader=[IO.StreamReader]::new($cmpEntry.Open())
   try {$cmpXml=[xml]$cmpReader.ReadToEnd()} finally {$cmpReader.Dispose()}
   foreach($cmpNode in $cmpXml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) {$cmpNode.InnerText=$cmpAuthor}
   foreach($cmpNode in $cmpXml.SelectNodes('//*[local-name()="Company" or local-name()="Manager"]')) {$cmpNode.InnerText=''}
   $cmpEntry.Delete()
   $cmpWriter=[IO.StreamWriter]::new($cmpArchive.CreateEntry($cmpName).Open(),[Text.UTF8Encoding]::new($false))
   try {$cmpWriter.Write($cmpXml.OuterXml)} finally {$cmpWriter.Dispose()}
  }
 } finally {$cmpArchive.Dispose()}
}
