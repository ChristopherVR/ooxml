# Synthetic references in an owned hidden instance; never change Office profile settings.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-review-inline-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$picturePath = Join-Path $destination 'reference.png'
[IO.File]::WriteAllBytes($picturePath, [Convert]::FromBase64String('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a5l8AAAAASUVORK5CYII='))
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
function Read-Reference($Document) {
    $types = @()
    for ($i = 1; $i -le $Document.Revisions.Count; $i++) { $types += [int]$Document.Revisions.Item($i).Type }
    return [ordered]@{ text = [string]$Document.Content.Text; revisions = [int]$Document.Revisions.Count; revisionTypes = $types; pictures = [int]$Document.InlineShapes.Count; footnotes = [int]$Document.Footnotes.Count; pages = [int]$Document.ComputeStatistics(2) }
}
function Add-Object($Document, [string]$Kind) {
    $range = $Document.Range(6, 6)
    switch ($Kind) {
        'break' { $range.InsertBreak(7) }
        'picture' {
            $picture = $Document.InlineShapes.AddPicture($picturePath, $false, $true, $range)
            $picture.Width = 24
            $picture.Height = 24
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($picture)
        }
        'note' {
            $note = $Document.Footnotes.Add($range, [Type]::Missing, 'Note text')
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($note)
        }
    }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
}
try {
    foreach ($kind in @('break', 'picture', 'note')) {
        foreach ($action in @('insert', 'delete')) {
            $name = "$kind-$action"
            $document = $application.Documents.Add()
            try {
                $document.TrackRevisions = $false
                $document.Content.Text = 'BeforeAfter'
                $document.Content.Font.Name = 'Arial'
                $document.Content.Font.Size = 12
                if ($action -eq 'delete') { Add-Object $document $kind }
                $before = Read-Reference $document
                $document.SaveAs2((Join-Path $destination "$name-before.docx"), 12)
                $document.TrackRevisions = $true
                if ($action -eq 'insert') { Add-Object $document $kind }
                else {
                    $position = if ($kind -eq 'break') { ([string]$document.Content.Text).IndexOf([char]12) } else { 6 }
                    if ($position -lt 0) { throw 'Expected a native page break' }
                    [void]$document.Range($position, $position + 1).Delete()
                }
                $tracked = Read-Reference $document
                $trackedPath = Join-Path $destination "$name-tracked.docx"
                $document.SaveAs2($trackedPath, 12)
                $document.Revisions.AcceptAll()
                $accepted = Read-Reference $document
                $document.SaveAs2((Join-Path $destination "$name-accepted.docx"), 12)
            } finally {
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
                $document = $null
            }
            $document = $application.Documents.Open($trackedPath, $false, $false)
            try {
                $document.Revisions.RejectAll()
                $rejected = Read-Reference $document
                $document.SaveAs2((Join-Path $destination "$name-rejected.docx"), 12)
                $cases += [ordered]@{ name = $name; before = $before; tracked = $tracked; accepted = $accepted; rejected = $rejected }
            } finally {
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
                $document = $null
            }
        }
    }
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; cases = $cases } |
        ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8NoBOM
    Write-Output $destination
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
