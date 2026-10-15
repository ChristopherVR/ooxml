# Author content controls of every common type in a separate hidden Word instance.
param([string]$OutputDirectory = (Join-Path $env:TEMP ('word-content-controls-' + [guid]::NewGuid())))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$destination = (Resolve-Path -LiteralPath $OutputDirectory).Path
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$document = $null
try {
    $document = $word.Documents.Add()
    $document.Content.Text = "Name: `rAgreed: `rColour: `rCity: `rDate: `rNotes`rAfter"
    function Range-Of($index) { $document.Paragraphs.Item($index).Range }
    function Inline-End($index) { $r = Range-Of $index; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0); $r }
    # 1 plain text, inline, with title, tag and a lock
    $text = $document.ContentControls.Add(1, (Inline-End 1))
    $text.Title = 'Full name'; $text.Tag = 'fullName'
    $text.LockContentControl = $true
    # 8 checkbox
    $check = $document.ContentControls.Add(8, (Inline-End 2))
    $check.Title = 'Agreed'; $check.Tag = 'agreed'; $check.Checked = $true
    # 4 drop-down list
    $drop = $document.ContentControls.Add(4, (Inline-End 3))
    $drop.Title = 'Colour'; $drop.Tag = 'colour'
    $drop.DropdownListEntries.Add('Red', 'r') | Out-Null
    $drop.DropdownListEntries.Add('Green', 'g') | Out-Null
    $drop.DropdownListEntries.Add('Blue', 'b') | Out-Null
    $drop.DropdownListEntries.Item(2).Select()
    # 3 combo box
    $combo = $document.ContentControls.Add(3, (Inline-End 4))
    $combo.Title = 'City'; $combo.Tag = 'city'
    $combo.DropdownListEntries.Add('Paris', 'p') | Out-Null
    $combo.DropdownListEntries.Add('Rome', 'r') | Out-Null
    # 6 date picker
    $date = $document.ContentControls.Add(6, (Inline-End 5))
    $date.Title = 'Start date'; $date.Tag = 'start'; $date.DateDisplayFormat = 'dd MMMM yyyy'
    # 0 rich text, block level, content locked
    $rich = $document.ContentControls.Add(0, (Range-Of 6))
    $rich.Title = 'Notes block'; $rich.Tag = 'notes'
    $rich.LockContents = $true
    $document.SaveAs2((Join-Path $destination 'content-controls.docx'), 12)
    $report = [ordered]@{ applicationVersion = [string]$word.Version; controls = $document.ContentControls.Count }
    $report | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'reference.json') -Encoding utf8
    $report | ConvertTo-Json
    $destination
} finally {
    if ($null -ne $document) { $document.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
}
