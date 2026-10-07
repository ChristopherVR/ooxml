# New documents only, in a separate hidden Word instance. PDF text is the page-field oracle.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-page-fields-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$wordPageFields = New-Object -ComObject Word.Application
$wordPageFields.Visible = $false
$wordPageFields.DisplayAlerts = 0
try {
    foreach ($kind in @('default', 'title-page', 'restart', 'even-odd', 'title-restart', 'even-restart', 'roman-even-restart', 'even-restart-odd', 'even-restart-one', 'even-page-restart', 'even-page-restart-odd', 'roman-even-page-restart')) {
        $document = $wordPageFields.Documents.Add()
        try {
            $beforeCount = if ($kind.Contains('even-page')) {56} else {2}
            $before = (1..$beforeCount | ForEach-Object { "Before$_`r" }) -join ''
            $document.Content.Text = $before + ((1..120 | ForEach-Object { "After$_`r" }) -join '')
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
            $document.Content.ParagraphFormat.LineSpacingRule = 4
            $document.Content.ParagraphFormat.LineSpacing = 12
            $document.Range($before.Length,$before.Length).InsertBreak(3)
            if ($kind.Contains('even')) { $document.PageSetup.OddAndEvenPagesHeaderFooter = -1 }
            if ($kind -like 'title*') { $document.Sections.Item(2).PageSetup.DifferentFirstPageHeaderFooter = -1 }
            foreach ($sectionIndex in @(1,2)) {
                $section = $document.Sections.Item($sectionIndex)
                $prefix = if ($sectionIndex -eq 1) {'A'} else {'B'}
                foreach ($slot in @(1,2,3)) {
                    $name = switch ($slot) { 1 {'default'} 2 {'first'} 3 {'even'} }
                    $header = $section.Headers.Item($slot)
                    $header.LinkToPrevious = $false
                    $header.Range.Text = "H-$prefix-$name"
                    $footer = $section.Footers.Item($slot)
                    $footer.LinkToPrevious = $false
                    $footer.Range.Text = "F-$prefix-$name P:"
                    $range = $footer.Range.Duplicate
                    $range.End = $range.End - 1
                    $range.Collapse(0)
                    $footer.Range.Fields.Add($range,-1,'PAGE',$true) | Out-Null
                    $range = $footer.Range.Duplicate
                    $range.End = $range.End - 1
                    $range.Collapse(0)
                    $range.InsertAfter(' S:')
                    $range = $footer.Range.Duplicate
                    $range.End = $range.End - 1
                    $range.Collapse(0)
                    $footer.Range.Fields.Add($range,-1,'SECTIONPAGES',$true) | Out-Null
                    if ($sectionIndex -eq 2 -and $kind.Contains('restart')) {
                        $footer.PageNumbers.RestartNumberingAtSection = $true
                        $footer.PageNumbers.StartingNumber = switch ($kind) { 'even-restart-odd' {11} 'even-page-restart-odd' {11} 'even-restart-one' {1} default {10} }
                        if ($kind.StartsWith('roman-')) { $footer.PageNumbers.NumberStyle = 2 }
                    }
                }
            }
            $document.Repaginate()
            foreach ($story in $document.StoryRanges) { $story.Fields.Update() | Out-Null }
            $document.SaveAs2((Join-Path $OutputDirectory "$kind.docx"),16)
            $document.ExportAsFixedFormat((Join-Path $OutputDirectory "$kind.pdf"),17)
        } finally {
            $document.Close(0)
            [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null
            $document = $null
        }
    }
    [ordered]@{application='Microsoft Word';version=$wordPageFields.Version;build=$wordPageFields.Build} |
        ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'application.json') -Encoding utf8
} finally {
    $wordPageFields.Quit(0)
    [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordPageFields) | Out-Null
}
Write-Output $OutputDirectory
