# Generate fixtures with bun scripts/generate-word-numbering-overrides.mjs <fixture-directory>.
# Record labels with an owned hidden Word application, opening synthetic sources read-only.
param([Parameter(Mandatory)][string]$FixtureDirectory, [Parameter(Mandatory)][string]$OutputPath, [ValidateSet('restart', 'start')][string]$Kind = 'restart')
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$referencePath = Join-Path $root "src/core/docx/__fixtures__/numbering-override-$Kind/native-reference.json"
$reference = Get-Content -LiteralPath $referencePath -Raw | ConvertFrom-Json
$fixtures = (Resolve-Path -LiteralPath $FixtureDirectory).Path
$application = $null
$document = $null
try {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    foreach ($case in $reference.cases) {
        try {
            $document = $application.Documents.Open((Join-Path $fixtures "$($case.name).docx"), $false, $true, $false)
            $labels = @()
            for ($index = 1; $index -le 8; $index++) {
                $paragraph = $document.Paragraphs.Item($index)
                $range = $paragraph.Range
                $format = $range.ListFormat
                try { $labels += [string]$format.ListString }
                finally {
                    foreach ($com in @($format, $range, $paragraph)) {
                        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)
                    }
                }
            }
            $case.labels = $labels
        } finally {
            if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null }
        }
    }
    $reference.wordVersion = [string]$application.Version
    $reference.wordBuild = [string]$application.Build
    $reference.recordedDate = (Get-Date).ToString('yyyy-MM-dd')
    $reference | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $OutputPath -Encoding utf8
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($application) { $application.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
}
