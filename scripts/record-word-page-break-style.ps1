# Create only synthetic documents in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-page-break-style-' + [guid]::NewGuid())))
$ErrorActionPreference='Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination=(Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference=New-Object -ComObject Word.Application
$wordReference.Visible=$false
$wordReference.DisplayAlerts=0
$document=$null
try {
 $document=$wordReference.Documents.Add()
 $document.Content.Text="First paragraph`rInherited page break`rExplicitly disabled"
 $document.Content.Font.Name='Arial'
 $document.Content.Font.Size=12
 $style=$document.Styles.Add('PageBreakBase',1)
 $style.ParagraphFormat.PageBreakBefore=-1
 $document.Paragraphs.Item(2).Range.Style=$style
 $document.Paragraphs.Item(3).Range.Style=$style
 $document.Paragraphs.Item(3).Range.ParagraphFormat.PageBreakBefore=0
 $document.Repaginate()
 $cases=@()
 for($index=1;$index -le 3;$index++) {
  $paragraph=$document.Paragraphs.Item($index)
  $cases += [ordered]@{text=[string]$paragraph.Range.Text;pageBreakBefore=[int]$paragraph.Range.ParagraphFormat.PageBreakBefore;page=[int]$paragraph.Range.Information(3)}
 }
 $document.SaveAs2((Join-Path $destination 'page-break-style.docx'),12)
 $report=[ordered]@{applicationVersion=[string]$wordReference.Version;applicationBuild=[string]$wordReference.Build;cases=$cases}
 $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
 $report | ConvertTo-Json -Depth 6
} finally {
 if($null -ne $document){$document.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)}
 $wordReference.Quit(0)
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
