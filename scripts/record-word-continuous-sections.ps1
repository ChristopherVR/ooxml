# Native Word reference, using only newly created documents in a separate hidden instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-continuous-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$cases = @()
try {
    foreach ($kind in @('same', 'left-margin', 'top-margin', 'columns', 'page-size', 'orientation')) {
        $document = $wordReference.Documents.Add()
        try {
            $document.Content.Text = "Before1`rBefore2`rAfter1`rAfter2`r"
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $document.PageSetup.TopMargin = 72
            $document.PageSetup.BottomMargin = 72
            $document.PageSetup.LeftMargin = 72
            $document.PageSetup.RightMargin = 72
            $document.Content.Font.Name = 'Arial'
            $document.Content.Font.Size = 12
            $document.Content.ParagraphFormat.SpaceBefore = 0
            $document.Content.ParagraphFormat.SpaceAfter = 0
            $document.Content.ParagraphFormat.LineSpacingRule = 4 # wdLineSpaceExactly
            $document.Content.ParagraphFormat.LineSpacing = 12
            $document.Range(16,16).InsertBreak(3) # wdSectionBreakContinuous
            $setup = $document.Sections.Item(2).PageSetup
            switch ($kind) {
                'left-margin' { $setup.LeftMargin = 108 }
                'top-margin' { $setup.TopMargin = 108 }
                'columns' { $setup.TextColumns.SetCount(2) }
                'page-size' { $setup.PageHeight = 720 }
                'orientation' { $setup.Orientation = 1 }
            }
            $document.ActiveWindow.View.Type = 3 # wdPrintView
            $document.Repaginate()
            $positions = @()
            foreach ($paragraph in $document.Paragraphs) {
                $range = $paragraph.Range.Duplicate
                $range.Collapse(1) # wdCollapseStart
                $positions += [ordered]@{
                    text=$paragraph.Range.Text.Trim()
                    section=$range.Information(2)
                    page=$range.Information(3)
                    xPt=$range.Information(5)
                    yPt=$range.Information(6)
                }
            }
            $document.SaveAs2((Join-Path $OutputDirectory "$kind.docx"),16)
            $document.ExportAsFixedFormat((Join-Path $OutputDirectory "$kind.pdf"),17)
            $cases += [ordered]@{name=$kind;pages=$document.ComputeStatistics(2);positions=$positions}
        } finally {
            $document.Close(0)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
            $document = $null
        }
    }
    [ordered]@{application='Microsoft Word';version=$wordReference.Version;build=$wordReference.Build;cases=$cases} |
        ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    $wordReference.Quit(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference) | Out-Null
}
Write-Output $OutputDirectory
