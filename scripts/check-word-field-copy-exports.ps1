# Compare only synthetic range-transfer exports in an owned hidden Word instance.
param([Parameter(Mandatory)][string]$ExportDirectory, [Parameter(Mandatory)][string]$ReferenceDirectory, [Parameter(Mandatory)][string]$ReportPath, [switch]$MoveResult)
$ErrorActionPreference = 'Stop'
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
$cases = @()
try {
    $exports = (Resolve-Path -LiteralPath $ExportDirectory).Path
    $references = (Resolve-Path -LiteralPath $ReferenceDirectory).Path
    foreach ($kind in @('whole', 'result', 'partial')) {
        if ($MoveResult -and $kind -eq 'whole') { continue }
        $reference = $null
        foreach ($directory in @($references, $exports)) {
            $document = $application.Documents.Open((Join-Path $directory "$kind.docx"), $false, $true, $false)
            $fieldIndex = if ($kind -eq 'whole') { 2 } else { 1 }
            $field = $document.Fields.Item($fieldIndex)
            $result = $field.Result
            $content = $document.Content
            $range = $null
            $font = $null
            try {
                $length = if ($kind -eq 'partial') { 1 } else { 5 }
                $range = if ($kind -eq 'whole') { $result.Duplicate } else { $document.Range($result.End + 1, $result.End + 1 + $length) }
                $font = $range.Font
                $entry = [ordered]@{ name = $kind; fields = [int]$document.Fields.Count; text = [string]$content.Text; copiedBold = [int]$font.Bold }
                if ($reference) {
                    if ($entry.fields -ne $reference.fields -or $entry.text -cne $reference.text -or $entry.copiedBold -ne $reference.copiedBold) {
                        throw "Native range-transfer comparison differs: $kind"
                    }
                    $cases += $entry
                }
                $reference = $entry
            } finally {
                foreach ($com in @($font, $range, $content, $result, $field)) {
                    if ($com) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) }
                }
                $document.Close(0)
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
                $document = $null
            }
        }
    }
    [ordered]@{ applicationVersion = [string]$application.Version; applicationBuild = [string]$application.Build; cases = $cases } |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $ReportPath -Encoding utf8NoBOM
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
