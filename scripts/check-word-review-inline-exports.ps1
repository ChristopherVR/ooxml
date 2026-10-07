# Reopen only synthetic exports in an owned hidden instance; leave their bytes unchanged.
param([Parameter(Mandatory)][string]$ExportDirectory, [Parameter(Mandatory)][string]$ReportPath, [switch]$IncludeStories, [switch]$RejectAll, [switch]$IncludeObjectFormatting)
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
            $beforeRevisions = [int]$document.Revisions.Count
            if ($RejectAll) { $document.RejectAllRevisions() }
            $noteRevisions = 0
            for ($i = 1; $i -le $document.Footnotes.Count; $i++) { $noteRevisions += [int]$document.Footnotes.Item($i).Range.Revisions.Count }
            $entry = [ordered]@{ name = $file.BaseName; text = [string]$document.Content.Text; revisions = [int]$document.Revisions.Count; pictures = [int]$document.InlineShapes.Count; footnotes = [int]$document.Footnotes.Count; footnoteRevisions = $noteRevisions; pages = [int]$document.ComputeStatistics(2) }
            if ($IncludeStories) { $entry.stories = @(Get-WordReviewStories $document) }
            if ($RejectAll) { $entry.beforeRevisions = $beforeRevisions }
            if ($IncludeObjectFormatting) {
                $objects = @()
                for ($i = 1; $i -le $document.InlineShapes.Count; $i++) { $objects += [ordered]@{ kind = 'picture'; bold = [int]$document.InlineShapes.Item($i).Range.Font.Bold } }
                for ($i = 1; $i -le $document.Footnotes.Count; $i++) { $objects += [ordered]@{ kind = 'note'; bold = [int]$document.Footnotes.Item($i).Reference.Font.Bold } }
                for ($i = 1; $i -le $document.Fields.Count; $i++) { $objects += [ordered]@{ kind = 'field'; bold = [int]$document.Fields.Item($i).Code.Font.Bold } }
                $breakPosition = ([string]$document.Content.Text).IndexOf([char]12)
                if ($breakPosition -ge 0) { $objects += [ordered]@{ kind = 'break'; bold = [int]$document.Range($breakPosition, $breakPosition + 1).Font.Bold } }
				$lineBreakPosition = ([string]$document.Content.Text).IndexOf([char]11)
				if ($lineBreakPosition -ge 0) { $objects += [ordered]@{ kind = 'line-break'; bold = [int]$document.Range($lineBreakPosition, $lineBreakPosition + 1).Font.Bold } }
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
