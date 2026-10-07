# Inspect only the supplied synthetic fixture in a separate hidden Word instance.
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [string]$OutputDirectory = (Join-Path $env:TEMP ('word-review-moves-' + [guid]::NewGuid()))
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $InputPath).Path
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$document = $null
function Read-Revisions($Document) {
    $entries = @()
    for ($i = 1; $i -le $Document.Revisions.Count; $i++) {
        $revision = $Document.Revisions.Item($i)
        $entries += [ordered]@{ type = [int]$revision.Type; text = [string]$revision.Range.Text; author = [string]$revision.Author }
    }
    return $entries
}
try {
    $document = $wordReference.Documents.Open($source, $false, $true, $false)
    $before = @(Read-Revisions $document)
    $names = @([regex]::Matches($document.WordOpenXML, 'w:name="(move[a-zA-Z0-9]+)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
    $document.SaveAs2((Join-Path $destination 'word-saved.docx'), 12)
    # Accept the first source/destination pair, leaving the second move untouched.
    $firstSource = $null
    for ($i = 1; $i -le $document.Revisions.Count; $i++) {
        $revision = $document.Revisions.Item($i)
        if ([int]$revision.Type -eq 14 -and $revision.Range.Text -like '*FirstMove*') { $firstSource = $revision; break }
    }
    if ($null -eq $firstSource) { throw 'Word did not recognize the first moved-from revision.' }
    $firstSource.Accept()
    $after = @(Read-Revisions $document)
    $textAfterAccept = [string]$document.Content.Text
    $document.SaveAs2((Join-Path $destination 'word-accepted-first.docx'), 12)
    $document.Close(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null
    $document = $wordReference.Documents.Open($source, $false, $true, $false)
    $firstSource = $null
    for ($i = 1; $i -le $document.Revisions.Count; $i++) {
        $revision = $document.Revisions.Item($i)
        if ([int]$revision.Type -eq 14 -and $revision.Range.Text -like '*FirstMove*') { $firstSource = $revision; break }
    }
    if ($null -eq $firstSource) { throw 'Word did not recognize the first moved-from revision.' }
    $firstSource.Reject()
    $afterReject = @(Read-Revisions $document)
    $document.SaveAs2((Join-Path $destination 'word-rejected-first.docx'), 12)
    $report = [ordered]@{
        applicationVersion = [string]$wordReference.Version
        applicationBuild = [string]$wordReference.Build
        names = $names
        before = $before
        afterAcceptFirst = $after
        textAfterAcceptFirst = $textAfterAccept
        afterRejectFirst = $afterReject
        textAfterRejectFirst = [string]$document.Content.Text
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 8
} finally {
    if ($null -ne $document) { $document.Close(0); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
