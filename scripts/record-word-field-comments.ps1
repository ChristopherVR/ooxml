# Record synthetic field comment scopes in an owned hidden Word instance.
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
    foreach ($kind in @('begin', 'result', 'whole')) {
        $document = $application.Documents.Open($source, $false, $true, $false)
        $field = $document.Fields.Item(1)
        $code = $field.Code
        $result = $field.Result
        try {
            if ($kind -eq 'begin') { $start = $code.Start - 1; $end = $code.Start }
            elseif ($kind -eq 'result') { $result.Text = 'ABCDE'; $start = $result.Start + 1; $end = $start + 1 }
            else { $start = $code.Start - 1; $end = $result.End + 1 }
            foreach ($author in @('Ada', 'Bob')) {
                $range = $document.Range($start, $end)
                $comment = $null
                try {
                    $comment = $document.Comments.Add($range, "$author native")
                    $comment.Author = $author
                } finally {
                    if ($comment) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comment) }
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
                }
            }
            $document.SaveAs2((Join-Path $output "$kind.docx"), 16)
        } finally {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($result)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($code)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($field)
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
