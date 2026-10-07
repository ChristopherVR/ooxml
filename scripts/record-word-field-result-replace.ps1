# Record replacements inside synthetic cached field results in an owned hidden Word instance.
param([Parameter(Mandatory)][string]$SourceDocument, [Parameter(Mandatory)][string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
$application = New-Object -ComObject Word.Application
$application.Visible = $false
$application.DisplayAlerts = 0
$document = $null
try {
    $source = (Resolve-Path -LiteralPath $SourceDocument).Path
    $output = [IO.Path]::GetFullPath($OutputDirectory)
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    foreach ($kind in @('partial', 'whole', 'start', 'end')) {
        $document = $application.Documents.Open($source, $false, $true, $false)
        $field = $document.Fields.Item(1)
        $result = $field.Result
        $font = $result.Font
        $range = $null
        try {
            $font.Bold = -1
            if ($kind -eq 'whole') { $start = $result.Start; $end = $result.End }
            elseif ($kind -eq 'start') { $start = $result.Start; $end = $start + 1 }
            elseif ($kind -eq 'end') { $end = $result.End; $start = $end - 1 }
            else { $start = $result.Start + 1; $end = $start + 1 }
            $range = $document.Range($start, $end)
            $range.Text = 'X'
            $currentResult = $field.Result
            try { Write-Output "$kind fields: $($document.Fields.Count); result: $($currentResult.Text)" }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($currentResult) }
            $document.SaveAs2((Join-Path $output "$kind.docx"), 16)
        } finally {
            foreach ($com in @($range, $font, $result, $field)) {
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
