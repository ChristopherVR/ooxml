# Reopen only synthetic exports in an owned hidden instance; leave their bytes unchanged.
param([Parameter(Mandatory)][string]$ExportDirectory, [Parameter(Mandatory)][string]$ReportPath, [switch]$IncludeStories)
$ErrorActionPreference = 'Stop'
if ($IncludeStories) { . (Join-Path $PSScriptRoot 'word-review-stories.ps1') }
$directory = (Resolve-Path -LiteralPath $ExportDirectory).Path
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
try {
    foreach ($file in (Get-ChildItem -LiteralPath $directory -Filter '*.docx' | Sort-Object Name)) {
        $document = $application.Documents.Open($file.FullName, $false, $true)
        try {
            $noteRevisions = 0
            for ($i = 1; $i -le $document.Footnotes.Count; $i++) { $noteRevisions += [int]$document.Footnotes.Item($i).Range.Revisions.Count }
            $entry = [ordered]@{ name = $file.BaseName; text = [string]$document.Content.Text; revisions = [int]$document.Revisions.Count; pictures = [int]$document.InlineShapes.Count; footnotes = [int]$document.Footnotes.Count; footnoteRevisions = $noteRevisions; pages = [int]$document.ComputeStatistics(2) }
            if ($IncludeStories) { $entry.stories = @(Get-WordReviewStories $document) }
            $cases += $entry
        } finally {
            $document.Close(0)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            $document = $null
        }
    }
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; cases = $cases } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding utf8NoBOM
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
