# Compare synthetic cached-result exports in an owned hidden Word instance.
param([Parameter(Mandatory)][string]$ExportDirectory, [Parameter(Mandatory)][string]$ReferenceDirectory, [Parameter(Mandatory)][string]$ReportPath)
$ErrorActionPreference = 'Stop'
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
try {
    $exports = (Resolve-Path -LiteralPath $ExportDirectory).Path
    $references = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
    foreach ($kind in @('partial', 'whole', 'start', 'end', 'empty', 'tracked-empty')) {
        $reference = $null
        $methods = if ($kind -in @('empty', 'tracked-empty')) { @('native', 'delete') } else { @('native', 'typing', 'paste') }
        foreach ($method in $methods) {
            $path = if ($method -eq 'native') { Join-Path $references "$kind.docx" } else { Join-Path $exports "$kind-$method.docx" }
            $document = $application.Documents.Open($path, $false, $true, $false)
            $content = $document.Content
            $fields = @()
            try {
                for ($index = 1; $index -le $document.Fields.Count; $index++) {
                    $field = $document.Fields.Item($index)
                    $result = $field.Result
                    $code = $field.Code
                    $font = $result.Font
                    try { $fields += [ordered]@{ result = [string]$result.Text; instruction = ([string]$code.Text).Trim(); bold = [int]$font.Bold } }
                    finally { foreach ($com in @($font, $code, $result, $field)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
                }
                $entry = [ordered]@{ name = "$kind-$method"; text = [string]$content.Text; fields = $fields; revisions = [int]$document.Revisions.Count }
                $snapshot = [ordered]@{ text = $entry.text; fields = $fields; revisions = $entry.revisions } | ConvertTo-Json -Depth 5 -Compress
                if ($reference -and $snapshot -cne $reference) { throw "Native cached-result comparison differs: $kind-$method`nExpected: $reference`nActual: $snapshot" }
                $reference = $snapshot
                $cases += $entry
            } finally {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($content)
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
                $document = $null
            }
        }
    }
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; cases = $cases } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding utf8NoBOM
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
