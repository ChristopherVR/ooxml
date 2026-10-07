# Reopen only synthetic exports in an owned hidden instance; leave their bytes unchanged.
param([Parameter(Mandatory)][string]$ExportDirectory, [Parameter(Mandatory)][string]$ReportPath, [switch]$IncludeStories, [switch]$RejectAll, [switch]$IncludeObjectFormatting, [switch]$IncludeAdvancedFormatting, [string]$FilePattern = '*.docx')
$ErrorActionPreference = 'Stop'
if ($IncludeStories) { . (Join-Path $PSScriptRoot 'word-review-stories.ps1') }
$directory = (Resolve-Path -LiteralPath $ExportDirectory).Path
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
function Get-WordObjectFormatting($kind, $range) {
    $font = $range.Font
    try {
        $entry = [ordered]@{ kind = $kind; bold = [int]$font.Bold }
        if ($IncludeAdvancedFormatting) {
            $entry.size = [double]$font.Size
            $entry.color = [int]$font.Color
            $entry.smallCaps = [int]$font.SmallCaps
            $entry.spacing = [double]$font.Spacing
            $entry.scale = [int]$font.Scaling
            $entry.position = [double]$font.Position
            $entry.kerning = [double]$font.Kerning
        }
        return $entry
    } finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($font)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
    }
}
try {
    foreach ($file in (Get-ChildItem -LiteralPath $directory -Filter $FilePattern | Sort-Object Name)) {
        $document = $application.Documents.Open($file.FullName, $false, $true)
        try {
            $beforeRevisions = [int]$document.Revisions.Count
            if ($RejectAll) { $document.RejectAllRevisions() }
            $noteRevisions = 0
            for ($i = 1; $i -le $document.Footnotes.Count; $i++) { $noteRevisions += [int]$document.Footnotes.Item($i).Range.Revisions.Count }
            $entry = [ordered]@{ name = $file.BaseName; text = [string]$document.Content.Text; revisions = [int]$document.Revisions.Count; pictures = [int]$document.InlineShapes.Count; footnotes = [int]$document.Footnotes.Count; footnoteRevisions = $noteRevisions; pages = [int]$document.ComputeStatistics(2) }
            if ($IncludeStories) { $entry.stories = @(Get-WordReviewStories $document) }
            if ($RejectAll) { $entry.beforeRevisions = $beforeRevisions }
            if ($IncludeObjectFormatting) {
                $objects = @()
                for ($i = 1; $i -le $document.InlineShapes.Count; $i++) { $objects += Get-WordObjectFormatting 'picture' $document.InlineShapes.Item($i).Range }
                for ($i = 1; $i -le $document.Footnotes.Count; $i++) { $objects += Get-WordObjectFormatting 'note' $document.Footnotes.Item($i).Reference }
                for ($i = 1; $i -le $document.Fields.Count; $i++) { $objects += Get-WordObjectFormatting 'field' $document.Fields.Item($i).Code }
                $breakPosition = ([string]$document.Content.Text).IndexOf([char]12)
                if ($breakPosition -ge 0) { $objects += Get-WordObjectFormatting 'break' ($document.Range($breakPosition, $breakPosition + 1)) }
				$lineBreakPosition = ([string]$document.Content.Text).IndexOf([char]11)
				if ($lineBreakPosition -ge 0) { $objects += Get-WordObjectFormatting 'line-break' ($document.Range($lineBreakPosition, $lineBreakPosition + 1)) }
                $entry.objects = $objects
            }
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
