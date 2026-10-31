# Creates a synthetic, formatted comment thread with Word 16, then anonymizes package properties.
param([switch]$Whitespace)
$ErrorActionPreference='Stop'
$parityFilename=if($Whitespace){'word-comment-whitespace-native.docx'}else{'word-comments-native.docx'}
$parityWord=New-Object -ComObject Word.Application
$parityWord.Visible=$false
$parityWord.DisplayAlerts=0
$parityDoc=$null
try {
 $parityDoc=$parityWord.Documents.Add()
 $parityDoc.Content.Text="Review this paragraph.`rAnother paragraph."
 $parityText=if($Whitespace){" lead `tA$([char]11)B$([char]11)C$([char]0x2011)D$([char]0xad)E`r end "}else{"Bold text and plain text.`rSecond paragraph."}
 $parityComment=$parityDoc.Comments.Add($parityDoc.Range(0,6),$parityText)
 $parityComment.Author='Parity Author'
 $parityComment.Initial='PA'
 $parityComment.Range.Words.Item(1).Bold=-1
 $parityComment.Range.Paragraphs.Item(2).Range.Font.Italic=-1
 $parityReply=$parityComment.Replies.Add($parityComment.Scope,'Reply text.')
 $parityReply.Author='Parity Reviewer'
 $parityReply.Initial='PR'
 $parityComment.Done=$true
 $parityDoc.SaveAs2((Join-Path $PSScriptRoot $parityFilename),16)
 Write-Output "Word $($parityWord.Version): $($parityDoc.Comments.Count) comments, $($parityComment.Replies.Count) replies"
} finally {
 if($parityDoc){$parityDoc.Close(0)}
 $parityWord.Quit()
 [Runtime.InteropServices.Marshal]::FinalReleaseComObject($parityWord)|Out-Null
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$parityArchive=[IO.Compression.ZipFile]::Open((Join-Path $PSScriptRoot $parityFilename),'Update')
try {
 foreach($parityName in @('docProps/core.xml','docProps/app.xml')) {
  $parityEntry=$parityArchive.GetEntry($parityName)
  if(-not $parityEntry){continue}
  $parityReader=[IO.StreamReader]::new($parityEntry.Open())
  try {$parityXml=[xml]$parityReader.ReadToEnd()} finally {$parityReader.Dispose()}
  foreach($parityNode in $parityXml.SelectNodes('//*[local-name()="creator" or local-name()="lastModifiedBy"]')) {$parityNode.InnerText='Parity Author'}
  foreach($parityNode in $parityXml.SelectNodes('//*[local-name()="Company" or local-name()="Manager"]')) {$parityNode.InnerText=''}
  $parityEntry.Delete()
  $parityWriter=[IO.StreamWriter]::new($parityArchive.CreateEntry($parityName).Open(),[Text.UTF8Encoding]::new($false))
  try {$parityWriter.Write($parityXml.OuterXml)} finally {$parityWriter.Dispose()}
 }
} finally {$parityArchive.Dispose()}
