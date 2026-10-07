# Record document-level review preferences without changing the user's Word profile.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-review-preferences-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$document = $null
try {
    $document = $wordReference.Documents.Add()
    $document.TrackRevisions = $false
    $document.Content.Text = 'Preference text'
    $document.Content.Font.Name = 'Arial'
    $document.Content.Font.Bold = 0
    $document.TrackFormatting = $false
    $document.TrackMoves = $false
    $document.TrackRevisions = $true
    $document.SaveAs2((Join-Path $destination 'preferences.docx'), 12)
    $document.Range(0, 10).Font.Bold = -1
    $document.Range(15, 15).InsertAfter('!')
    $document.SaveAs2((Join-Path $destination 'edited.docx'), 12)
    $revisions = @()
    foreach ($revision in $document.Revisions) {
        $revisions += [ordered]@{ type = [int]$revision.Type; text = [string]$revision.Range.Text }
    }
    $document.RejectAllRevisions()
    $report = [ordered]@{
        applicationVersion = [string]$wordReference.Version
        applicationBuild = [string]$wordReference.Build
        trackFormatting = [bool]$document.TrackFormatting
        trackMoves = [bool]$document.TrackMoves
        revisions = $revisions
        rejectedText = [string]$document.Content.Text
        rejectedFirstWordBold = [int]$document.Range(0, 10).Font.Bold
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 5
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
