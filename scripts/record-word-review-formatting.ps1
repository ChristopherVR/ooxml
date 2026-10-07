# Create only synthetic documents in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-review-formatting-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$document = $null
$cases = @()
function Read-Formatting($Document) {
    $range = $Document.Range(0, 9)
    return [ordered]@{ bold = [int]$range.Font.Bold; italic = [int]$range.Font.Italic; size = [double]$range.Font.Size; color = [int]$range.Font.Color; text = [string]$range.Text }
}
try {
    foreach ($name in @('bold', 'multiple')) {
        $document = $wordReference.Documents.Add()
        try {
            $document.TrackRevisions = $false
            $document.Content.Text = 'Format me'
            $range = $document.Range(0, 9)
            $range.Font.Name = 'Arial'
            $range.Font.Size = 12
            $range.Font.Bold = 0
            $range.Font.Italic = 0
            if ($name -eq 'multiple') { $range.Font.Bold = -1; $range.Font.Color = 16711680 }
            $before = Read-Formatting $document
            $document.SaveAs2((Join-Path $destination "$name-before.docx"), 12)
            $document.TrackRevisions = $true
            if ($name -eq 'bold') { $range.Font.Bold = -1 }
            else { $range.Font.Bold = 0; $range.Font.Italic = -1; $range.Font.Size = 20; $range.Font.Color = 32768 }
            $tracked = Read-Formatting $document
            $revisionTypes = @()
            for ($i = 1; $i -le $document.Revisions.Count; $i++) { $revisionTypes += [int]$document.Revisions.Item($i).Type }
            $document.SaveAs2((Join-Path $destination "$name-tracked.docx"), 12)
            $document.Revisions.RejectAll()
            $rejected = Read-Formatting $document
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
} finally {
    if ($null -ne $document) { $document.Close(0); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
