# Native synthetic references, using one owned hidden instance and no profile changes.
param([Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'word-review-stories.ps1')
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
try {
    $document = $application.Documents.Add()
    $document.TrackRevisions = $false
    $document.Content.Text = 'Body text'
    $document.Content.Font.Name = 'Arial'
    $document.Content.Font.Size = 12
    $document.Sections.Item(1).Headers.Item(1).Range.Text = 'Header text'
    $document.Sections.Item(1).Footers.Item(1).Range.Text = 'Footer text'
    [void]$document.Footnotes.Add($document.Range(4, 4), [Type]::Missing, 'Footnote text')
    $end = [int]$document.Content.End - 1
    [void]$document.Endnotes.Add($document.Range($end, $end), [Type]::Missing, 'Endnote text')
    $before = Get-WordReviewStories $document
    $document.SaveAs2((Join-Path $destination 'all-stories-before.docx'), 12)
    $document.TrackRevisions = $true
    $document.Range(0, 4).Font.Italic = -1
    $ranges = @($document.Sections.Item(1).Headers.Item(1).Range, $document.Sections.Item(1).Footers.Item(1).Range, $document.Footnotes.Item(1).Range, $document.Endnotes.Item(1).Range)
    $targets = @('Header text', 'Footer text', 'Footnote text', 'Endnote text')
    for ($i = 0; $i -lt $ranges.Count; $i++) {
        $range = $ranges[$i]
        $offset = ([string]$range.Text).IndexOf($targets[$i], [StringComparison]::Ordinal)
        if ($offset -lt 0) { throw "Expected native story text $($targets[$i])" }
        $start = [int]$range.Start + $offset
        $range.SetRange($start, $start + $targets[$i].Length)
        $range.Font.Bold = -1
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
    }
    $tracked = Get-WordReviewStories $document
    $trackedPath = Join-Path $destination 'all-stories-tracked.docx'
    $document.SaveAs2($trackedPath, 12)
    $document.AcceptAllRevisions()
    $accepted = Get-WordReviewStories $document
    $document.SaveAs2((Join-Path $destination 'all-stories-accepted.docx'), 12)
    $document.Close(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $application.Documents.Open($trackedPath, $false, $false)
    $document.RejectAllRevisions()
    $rejected = Get-WordReviewStories $document
    $document.SaveAs2((Join-Path $destination 'all-stories-rejected.docx'), 12)
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; before = $before; tracked = $tracked; accepted = $accepted; rejected = $rejected } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8NoBOM
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
