# Native object run-property references in an owned hidden instance.
param([Parameter(Mandatory)][string]$OutputDirectory, [ValidateSet('picture', 'note', 'break', 'field', 'line-break')][string[]]$Kinds = @('picture', 'note', 'break', 'field'))
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'word-review-stories.ps1')
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$picturePath = Join-Path $destination 'reference.png'
if ($Kinds -contains 'picture') { [IO.File]::WriteAllBytes($picturePath, [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a5l8AAAAASUVORK5CYII=')) }
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
try {
    foreach ($kind in $Kinds) {
        $document = $application.Documents.Add()
        $document.TrackRevisions = $false
        $document.Content.Text = 'BeforeAfter'
        $document.Content.Font.Name = 'Arial'
        $document.Content.Font.Size = 12
        $range = $document.Range(6, 6)
        switch ($kind) {
            'picture' { [void]$document.InlineShapes.AddPicture($picturePath, $false, $true, $range) }
            'note' { [void]$document.Footnotes.Add($range, [Type]::Missing, 'Note text') }
            'break' { $range.InsertBreak(7) }
            'line-break' { $range.InsertBreak(6) }
            'field' { [void]$document.Fields.Add($range, 33, [Type]::Missing, $false) }
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
        $before = Get-WordReviewStories $document
        $document.SaveAs2((Join-Path $destination "$kind-before.docx"), 12)
        $document.TrackRevisions = $true
        $range = switch ($kind) {
            'picture' { $document.InlineShapes.Item(1).Range }
            'note' { $document.Footnotes.Item(1).Reference }
            'break' { $position = ([string]$document.Content.Text).IndexOf([char]12); $document.Range($position, $position + 1) }
            'line-break' { $position = ([string]$document.Content.Text).IndexOf([char]11); $document.Range($position, $position + 1) }
            'field' { $document.Fields.Item(1).Code }
        }
        $range.Font.Bold = -1
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
        $tracked = Get-WordReviewStories $document
        $trackedPath = Join-Path $destination "$kind-tracked.docx"
        $document.SaveAs2($trackedPath, 12)
        $document.AcceptAllRevisions()
        $accepted = Get-WordReviewStories $document
        $document.SaveAs2((Join-Path $destination "$kind-accepted.docx"), 12)
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        $document = $application.Documents.Open($trackedPath, $false, $false)
        $document.RejectAllRevisions()
        $rejected = Get-WordReviewStories $document
        $document.SaveAs2((Join-Path $destination "$kind-rejected.docx"), 12)
        $cases += [ordered]@{ name = $kind; before = $before; tracked = $tracked; accepted = $accepted; rejected = $rejected }
        $document.Close(0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        $document = $null
    }
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; cases = $cases } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8NoBOM
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
