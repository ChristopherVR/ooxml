param([string]$Out)
Add-Type -AssemblyName System.Windows.Forms
$word = New-Object -ComObject Word.Application
$word.Visible = $false
function Grab($sel, $name) {
  $sel.Copy()
  Start-Sleep -Milliseconds 500
  $data = [System.Windows.Forms.Clipboard]::GetData('HTML Format')
  $data = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding(1252).GetBytes([string]$data))
  # Word links its temporary clip files under the user's profile; keep no local path.
  $data = $data -replace 'file:///[^"\s>]*?/msohtmlclip\d*/\d+/', 'file:///C:/Temp/msohtmlclip1/01/'
  [System.IO.File]::WriteAllText((Join-Path $Out "$name.html"), [string]$data, (New-Object System.Text.UTF8Encoding($false)))
}
try {
  # 1. styles and run formatting
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.Style = $doc.Styles.Item(-2); $sel.TypeText('Clipboard Heading'); $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-1)
  $sel.TypeText('Plain then ')
  $sel.Font.Bold = 1; $sel.TypeText('bold'); $sel.Font.Bold = 0
  $sel.TypeText(' and ')
  $sel.Font.SmallCaps = 1; $sel.TypeText('small caps'); $sel.Font.SmallCaps = 0
  $sel.TypeText(' and ')
  $s = $sel.Range.Start; $sel.TypeText('highlighted'); $doc.Range($s, $sel.Range.Start).HighlightColorIndex = 7
  $sel.TypeText(' and ')
  $sel.Font.DoubleStrikeThrough = 1; $sel.TypeText('double struck'); $sel.Font.DoubleStrikeThrough = 0
  $sel.TypeText(' and ')
  $sel.Font.Underline = 3; $sel.TypeText('double underline'); $sel.Font.Underline = 0
  $sel.TypeText(' and ')
  $sel.Font.Color = 255; $sel.TypeText('red'); $sel.Font.Color = -16777216
  $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-181); $sel.TypeText('A quoted paragraph.'); $sel.TypeParagraph()
  $sel.Style = $doc.Styles.Item(-3); $sel.TypeText('Second level heading'); $sel.TypeParagraph()
  $cs = $doc.Styles.Add('Callout Text', 1); $cs.Font.Italic = 1; $cs.Font.Color = 12611584
  $sel.Style = $cs; $sel.TypeText('Custom styled')
  $doc.Content.Select(); Grab $sel 'word-styles'
  $doc.Close(0)

  # 2. bullet multi-level list
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.TypeText('Before list'); $sel.TypeParagraph()
  $sel.TypeText('First bullet'); $sel.TypeParagraph()
  $sel.TypeText('Nested bullet'); $sel.TypeParagraph()
  $sel.TypeText('Deeper bullet'); $sel.TypeParagraph()
  $sel.TypeText('Back to top'); $sel.TypeParagraph()
  $sel.TypeText('After list')
  $r = $doc.Range($doc.Paragraphs.Item(2).Range.Start, $doc.Paragraphs.Item(5).Range.End)
  $tmpl = $doc.ListTemplates.Add($true)
  $glyphs = @([string][char]0xF0B7, 'o', [string][char]0xF0A7)
  $fonts = @('Symbol', 'Courier New', 'Wingdings')
  for ($i = 1; $i -le 3; $i++) {
    $lv = $tmpl.ListLevels.Item($i)
    $lv.NumberFormat = $glyphs[$i - 1]
    $lv.Font.Name = $fonts[$i - 1]
    $lv.NumberStyle = 23
    $lv.NumberPosition = 18 + 18 * ($i - 1)
    $lv.TextPosition = 36 + 18 * ($i - 1)
  }
  $r.ListFormat.ApplyListTemplate($tmpl, $false, 0)
  $doc.Paragraphs.Item(3).Range.ListFormat.ListIndent()
  $doc.Paragraphs.Item(4).Range.ListFormat.ListIndent(); $doc.Paragraphs.Item(4).Range.ListFormat.ListIndent()
  $doc.Content.Select(); Grab $sel 'word-bullets'
  $doc.Close(0)

  # 3. multilevel numbered list
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.TypeText('Step one'); $sel.TypeParagraph()
  $sel.TypeText('Sub step a'); $sel.TypeParagraph()
  $sel.TypeText('Sub sub step'); $sel.TypeParagraph()
  $sel.TypeText('Sub step b'); $sel.TypeParagraph()
  $sel.TypeText('Step two')
  $tmpl = $word.ListGalleries.Item(3).ListTemplates.Item(1)
  $doc.Content.ListFormat.ApplyListTemplate($tmpl, $false, 0)
  $doc.Paragraphs.Item(2).Range.ListFormat.ListIndent()
  $doc.Paragraphs.Item(3).Range.ListFormat.ListIndent(); $doc.Paragraphs.Item(3).Range.ListFormat.ListIndent()
  $doc.Paragraphs.Item(4).Range.ListFormat.ListIndent()
  $doc.Content.Select(); Grab $sel 'word-numbered'
  $doc.Close(0)

  # 4. footnote, field and table
  $doc = $word.Documents.Add()
  $sel = $word.Selection
  $sel.TypeText('Text with a note')
  $fn = $doc.Footnotes.Add($sel.Range)
  $fn.Range.Text = 'The footnote text.'
  $doc.Content.InsertParagraphAfter()
  $end = $doc.Content.End - 1
  $r = $doc.Range($end, $end); $r.InsertAfter('Page ')
  $end = $doc.Content.End - 1
  $doc.Fields.Add($doc.Range($end, $end), 33, '', $false) | Out-Null
  $doc.Content.InsertParagraphAfter()
  $end = $doc.Content.End - 1
  $t = $doc.Tables.Add($doc.Range($end, $end), 2, 2)
  $t.Borders.Enable = 1
  $t.Cell(1,1).Range.Text = 'A1'; $t.Cell(1,2).Range.Text = 'B1'
  $t.Cell(2,1).Range.Text = 'A2'; $t.Cell(2,2).Range.Text = 'B2'
  $t.Cell(1,1).Range.Font.Bold = 1
  $doc.Content.Select(); Grab $sel 'word-table-notes'
  $doc.Close(0)
} finally { $word.Quit() }
