# Record synthetic source formatting in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-opaque-properties-' + [guid]::NewGuid())))
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
    $document.Content.Text = 'Outlined text'
    $document.Content.Font.Name = 'Arial'
    $document.Content.Font.Size = 12
    $document.Content.Font.Outline = -1
    $document.Content.Font.Shadow = -1
    $document.SaveAs2((Join-Path $destination 'outline-shadow.docx'), 12)
    $report = [ordered]@{
        applicationVersion = [string]$wordReference.Version
        applicationBuild = [string]$wordReference.Build
        text = [string]$document.Content.Text
        outline = [int]$document.Content.Font.Outline
        shadow = [int]$document.Content.Font.Shadow
        fontFamily = [string]$document.Content.Font.Name
        fontSize = [double]$document.Content.Font.Size
    }
    $report | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json -Depth 5
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $wordReference.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference)
}
