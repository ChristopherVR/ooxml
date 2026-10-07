# Native reference: fresh documents in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-continuous-tables-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$wordReference = New-Object -ComObject Word.Application
$wordReference.Visible = $false
$wordReference.DisplayAlerts = 0
$cases = @()
try {
    foreach ($kind in @('table-even', 'table-odd', 'table-overflow')) {
        $document = $wordReference.Documents.Add()
        try {
            $count = switch ($kind) { 'table-even' { 4 }; 'table-odd' { 5 }; 'table-overflow' { 120 } }
            $document.Content.Text = "After1`rAfter2`r"
            $document.PageSetup.PageWidth = 612
            $document.PageSetup.PageHeight = 792
            $document.PageSetup.TopMargin = 72
            $document.PageSetup.BottomMargin = 72
            $document.PageSetup.LeftMargin = 72
            $document.PageSetup.RightMargin = 72
            $table = $document.Tables.Add($document.Range(0,0),$count,1)
            $table.AllowAutoFit = $false
            $table.PreferredWidthType = 3 # wdPreferredWidthPoints
            $table.PreferredWidth = 216
            $table.Columns.Item(1).Width = 216
            $table.TopPadding = 0
            $table.BottomPadding = 0
            $table.LeftPadding = 0
            $table.RightPadding = 0
            $table.Rows.HeightRule = 2 # wdRowHeightExactly
            $table.Rows.Height = 12
            $table.Rows.AllowBreakAcrossPages = 0
            $table.Borders.Enable = 0
            for ($row=1; $row -le $count; $row++) { $table.Cell($row,1).Range.Text = "Row$row" }
            $document.Content.Font.Name = 'Arial'
            $document.Content.Font.Size = 12
            $document.Content.ParagraphFormat.SpaceBefore = 0
            $document.Content.ParagraphFormat.SpaceAfter = 0
            $document.Content.ParagraphFormat.LineSpacingRule = 4
            $document.Content.ParagraphFormat.LineSpacing = 12
            $end = $table.Range.End
            $document.Range($end,$end).InsertBreak(3) # wdSectionBreakContinuous
            $document.Sections.Item(1).PageSetup.TextColumns.SetCount(2)
            $document.Sections.Item(2).PageSetup.TextColumns.SetCount(1)
            $document.ActiveWindow.View.Type = 3
            $document.Repaginate()
            $positions = @()
            foreach ($paragraph in $document.Paragraphs) {
                $range = $paragraph.Range.Duplicate
                $range.Collapse(1)
                $positions += [ordered]@{
                    text=$paragraph.Range.Text.Trim([char[]]@([char]13,[char]7,[char]12))
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
        }
    }
    [ordered]@{application='Microsoft Word';version=$wordReference.Version;build=$wordReference.Build;cases=$cases} |
        ConvertTo-Json -Depth 8 | Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    $wordReference.Quit(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordReference) | Out-Null
}
Write-Output $OutputDirectory
