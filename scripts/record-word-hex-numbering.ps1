# Open only the generated synthetic package in an owned hidden Word process.
param([Parameter(Mandatory)][string]$FixturePath, [Parameter(Mandatory)][string]$OutputPath)
$ErrorActionPreference = 'Stop'
$fixture = (Resolve-Path -LiteralPath $FixturePath).Path
$zip = [System.IO.Compression.ZipFile]::OpenRead($fixture)
try {
    $reader = [System.IO.StreamReader]::new($zip.GetEntry('word/numbering.xml').Open())
    try { [xml]$numbering = $reader.ReadToEnd() } finally { $reader.Dispose() }
    $templates = @($numbering.GetElementsByTagName('lvlText', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') | ForEach-Object { $_.GetAttribute('val', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main') })
} finally { $zip.Dispose() }
$application = $null
$document = $null
try {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $document = $application.Documents.Open($fixture, $false, $true, $false)
    $cases = @()
    for ($index = 1; $index -le $templates.Count; $index++) {
        $paragraph = $document.Paragraphs.Item($index)
        $range = $paragraph.Range
        $format = $range.ListFormat
        try {
            $template = $templates[$index - 1]
            $name = switch ($template) { '%1' { 'raw' } '%1.' { 'dot' } '[%1]' { 'bracket' } default { throw "Unexpected fixture template: $template" } }
            $cases += @{ name = "$name-$([int]$format.ListValue)"; template = $template; value = [int]$format.ListValue; label = [string]$format.ListString }
        } finally {
            foreach ($com in @($format, $range, $paragraph)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)
            }
        }
    }
    @{
        wordVersion = [string]$application.Version
        wordBuild = [string]$application.Build
        recordedDate = (Get-Date).ToString('yyyy-MM-dd')
        scope = 'installed perpetual Word; not current Microsoft 365 subscription certification'
        source = 'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c87674b0-833c-4019-a21d-9f99dac53178'
        generator = 'src/core/docx/test-support/hex-numbering-fixture.ts: hexNumberingFixture()'
        recorder = 'scripts/record-word-hex-numbering.ps1'
        fixtureScope = 'synthetic nineteen independent single-level hex lists; raw placeholder, dot suffix and bracket controls; includes zero and 16-bit boundaries; no user documents'
        cases = $cases
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding utf8
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($application) { $application.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
}
