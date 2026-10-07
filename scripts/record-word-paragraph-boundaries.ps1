param([Parameter(Mandatory)][string]$FixtureDirectory, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$fixtures = (Resolve-Path -LiteralPath $FixtureDirectory).Path
$output = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($output) | Out-Null
$application = $null
$document = $null
try {
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $cases = @()
    foreach ($action in @('split', 'join')) {
        try {
            $document = $application.Documents.Open((Join-Path $fixtures "$action.docx"), $false, $true, $false)
            $document.TrackRevisions = $true
            $range = $document.Range(2, $(if ($action -eq 'split') { 2 } else { 3 }))
            try {
                if ($action -eq 'split') { $range.Text = "`r" } else { [void]$range.Delete() }
            } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range) }
            $paragraphs = @()
            for ($index = 1; $index -le $document.Paragraphs.Count; $index++) {
                $paragraph = $document.Paragraphs.Item($index)
                $range = $paragraph.Range
                $format = $paragraph.Format
                try { $paragraphs += @{ text = ([string]$range.Text).TrimEnd("`r"); alignment = [int]$format.Alignment } }
                finally { foreach ($com in @($format, $range, $paragraph)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
            }
            $document.SaveAs2((Join-Path $output "$action.docx"), 12)
            $cases += @{ action = $action; paragraphs = $paragraphs }
        } finally {
            if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null }
        }
    }
    @{ wordVersion = [string]$application.Version; wordBuild = [string]$application.Build; recordedDate = (Get-Date).ToString('yyyy-MM-dd'); scope = 'installed perpetual Word Range.Text CR and Range.Delete APIs, not interactive Enter or current Microsoft 365 certification'; generator = 'src/core/docx/test-support/paragraph-boundary-fixture.ts'; recorder = 'scripts/record-word-paragraph-boundaries.ps1'; cases = $cases } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $output 'native-reference.json') -Encoding utf8
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    if ($application) { $application.Quit(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) }
}
