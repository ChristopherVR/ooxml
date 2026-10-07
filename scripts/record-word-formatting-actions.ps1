# Record formatting actions in synthetic documents owned by this process.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-formatting-actions-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$document = $null
$cases = @()
try {
    foreach ($name in @('bold', 'bold-italic', 'bold-off', 'own-insertion')) {
        $document = $wordReference.Documents.Add()
        $document.TrackRevisions = $false
        $document.Content.Text = 'Format me'
        $document.Content.Font.Name = 'Arial'
        $document.Content.Font.Size = 12
        $document.Content.Font.Bold = 0
        $document.Content.Font.Italic = 0
        if ($name -eq 'bold') { $document.SaveAs2((Join-Path $destination 'baseline.docx'), 12) }
        $document.TrackFormatting = $true
        $document.TrackRevisions = $true
        if ($name -eq 'own-insertion') { $document.Range(9, 9).InsertAfter('!') }
        $range = if ($name -eq 'own-insertion') { $document.Range(9, 10) } else { $document.Range(0, 6) }
        $range.Font.Bold = -1
        if ($name -eq 'bold-italic') { $range.Font.Italic = -1 }
        if ($name -eq 'bold-off') { $range.Font.Bold = 0 }
        $document.SaveAs2((Join-Path $destination ($name + '.docx')), 12)
        $revisions = @()
        foreach ($revision in $document.Revisions) {
            $revisions += [ordered]@{ type = [int]$revision.Type; author = [string]$revision.Author; text = [string]$revision.Range.Text }
        }
        $document.RejectAllRevisions()
        $cases += [ordered]@{ name = $name; revisions = $revisions; rejectedText = [string]$document.Content.Text; rejectedBold = [int]$document.Content.Font.Bold; rejectedItalic = [int]$document.Content.Font.Italic }
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        $document = $null
    }
    $report = [ordered]@{ applicationVersion = [string]$wordReference.Version; applicationBuild = [string]$wordReference.Build; cases = $cases }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 8
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
