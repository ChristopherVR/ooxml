# Create only synthetic documents in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-review-paragraph-formatting-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$document = $null
$cases = @()
function Read-ParagraphFormatting($Document) {
    $format = $Document.Paragraphs.Item(1).Range.ParagraphFormat
    return [ordered]@{ alignment = [int]$format.Alignment; before = [double]$format.SpaceBefore; after = [double]$format.SpaceAfter; left = [double]$format.LeftIndent; first = [double]$format.FirstLineIndent; line = [double]$format.LineSpacing; lineRule = [int]$format.LineSpacingRule; keepNext = [int]$format.KeepWithNext }
}
try {
    foreach ($name in @('alignment', 'spacing', 'indent', 'multiple')) {
        $document = $wordReference.Documents.Add()
        try {
            $document.TrackRevisions = $false
            $document.Content.Text = 'Paragraph formatting'
            $document.Content.Font.Name = 'Arial'
            $document.Content.Font.Size = 12
            $format = $document.Paragraphs.Item(1).Range.ParagraphFormat
            $format.Alignment = 0
            $format.SpaceBefore = 0
            $format.SpaceAfter = 0
            $format.LeftIndent = 0
            $format.FirstLineIndent = 0
            $format.LineSpacingRule = 4
            $format.LineSpacing = 12
            $format.KeepWithNext = 0
            $before = Read-ParagraphFormatting $document
            $document.SaveAs2((Join-Path $destination "$name-before.docx"), 12)
            $document.TrackRevisions = $true
            if ($name -in @('alignment', 'multiple')) { $format.Alignment = 1 }
            if ($name -in @('spacing', 'multiple')) { $format.SpaceBefore = 18; $format.SpaceAfter = 6; $format.LineSpacing = 24 }
            if ($name -in @('indent', 'multiple')) { $format.LeftIndent = 36; $format.FirstLineIndent = 18 }
            if ($name -eq 'multiple') { $format.KeepWithNext = -1 }
            $tracked = Read-ParagraphFormatting $document
            $revisionTypes = @()
            for ($i = 1; $i -le $document.Revisions.Count; $i++) { $revisionTypes += [int]$document.Revisions.Item($i).Type }
            $document.SaveAs2((Join-Path $destination "$name-tracked.docx"), 12)
            $document.Revisions.RejectAll()
            $rejected = Read-ParagraphFormatting $document
            $document.SaveAs2((Join-Path $destination "$name-rejected.docx"), 12)
            $cases += [ordered]@{ name = $name; before = $before; tracked = $tracked; rejected = $rejected; revisionTypes = $revisionTypes }
        } finally {
            $document.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            $document = $null
        }
    }
    $report = [ordered]@{ applicationVersion = [string]$wordReference.Version; applicationBuild = [string]$wordReference.Build; cases = $cases }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 8
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
