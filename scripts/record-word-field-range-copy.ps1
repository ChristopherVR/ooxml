# Transfer synthetic ranges in an owned hidden Word instance; never change the source file.
param([Parameter(Mandatory)][string]$SourceDocument, [Parameter(Mandatory)][string]$OutputDirectory, [switch]$MoveResult)
$ErrorActionPreference = 'Stop'
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
try {
    $source = (Resolve-Path -LiteralPath $SourceDocument).Path
    $output = [IO.Path]::GetFullPath($OutputDirectory)
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    foreach ($kind in @('whole', 'result', 'partial')) {
        if ($MoveResult -and $kind -eq 'whole') { continue }
        $document = $application.Documents.Open($source, $false, $true, $false)
        $field = $document.Fields.Item(1)
        $code = $field.Code
        $result = $field.Result
        $font = $result.Font
        $range = $null
        $target = $null
        $formatted = $null
        try {
            $font.Bold = -1
            if ($kind -eq 'whole') { $start = $code.Start - 1; $end = $result.End + 1 }
            elseif ($kind -eq 'result') { $start = $result.Start; $end = $result.End }
            else { $start = $result.Start + 1; $end = $start + 1 }
            $range = $document.Range($start, $end)
            $target = $document.Range($result.End + 1, $result.End + 1)
            $formatted = $range.FormattedText
            $target.FormattedText = $formatted
            if ($MoveResult) { [void]$range.Delete() }
            Write-Output "$kind fields: $($document.Fields.Count); text: $($document.Content.Text)"
            $document.SaveAs2((Join-Path $output "$kind.docx"), 16)
        } finally {
            foreach ($com in @($formatted, $target, $range, $font, $result, $code, $field)) {
                if ($com) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) }
            }
            $document.Close(0)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
            $document = $null
        }
    }
} finally {
    if ($document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $application.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
