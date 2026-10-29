# Regenerates the Word reference corpus (manual; needs desktop Word on Windows, not run in CI).
#
#   pwsh -File src/core/scripts/docx-word-corpus.ps1 [-OutDir <folder>]
#
# Word authors ~12 small documents, then for each one records, into <OutDir> (default
# src/core/docx/__fixtures__/word-corpus):
#   <name>.docx        the document as Word saved it
#   <name>.facts.json  page count, per-paragraph start page and line-start character offsets, floating
#                      shape positions, Word version and fonts (written by this script, from Range.Information)
#   <name>-p<N>.png    each page as Word draws it (Pages(N).EnhMetaFileBits), rasterised at $dpi
# Every body font is Courier New so a headless test can measure text exactly (advance 0.6 em);
# this keeps the line-break comparison independent of host font availability. See
# viewers/docx/docs/verification.md.
param(
	[string]$OutDir = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/word-corpus'),
	[int]$Dpi = 50,
	[int]$WaitMinutes = 15,
	# Regenerate only these documents (names without extension); default is the whole corpus.
	[string[]]$Only = @()
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force $OutDir | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-corpus-work'
New-Item -ItemType Directory -Force $work | Out-Null

# Never share a Word that someone else is using.
for ($waited = 0; (Get-Process WINWORD -ErrorAction SilentlyContinue); $waited += 30) {
	if ($waited -ge $WaitMinutes * 60) { throw 'WINWORD.EXE kept running; giving up.' }
	Start-Sleep -Seconds 30
}

$font = 'Courier New'
$sentence = 'The quick brown fox jumps over the lazy dog while seven wizards box quickly. '
function Body([int]$n) { ($sentence * $n).TrimEnd() }

# Test picture: 160x100 gradient with a frame, written once.
$png = Join-Path $work 'picture.png'
$bmp = New-Object Drawing.Bitmap 160, 100
$g = [Drawing.Graphics]::FromImage($bmp)
$brush = New-Object Drawing.Drawing2D.LinearGradientBrush ([Drawing.Point]::new(0, 0)), ([Drawing.Point]::new(160, 100)), ([Drawing.Color]::SteelBlue), ([Drawing.Color]::Orange)
$g.FillRectangle($brush, 0, 0, 160, 100)
$g.DrawRectangle([Drawing.Pens]::Black, 0, 0, 159, 99)
$g.Dispose(); $bmp.Save($png, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()

function Initialize-Doc($word) {
	$doc = $word.Documents.Add()
	$ps = $doc.PageSetup
	$ps.PageWidth = 612; $ps.PageHeight = 792
	$ps.TopMargin = 72; $ps.BottomMargin = 72; $ps.LeftMargin = 72; $ps.RightMargin = 72
	foreach ($s in @('Normal', 'Heading 1', 'Heading 2', 'Title', 'List Paragraph', 'List Bullet', 'List Number', 'Header', 'Footer', 'Footnote Text', 'Endnote Text', 'TOC 1', 'TOC 2', 'Caption', 'Comment Text')) {
		try {
			$st = $doc.Styles.Item($s)
			$st.Font.Name = $font
			if ($s -eq 'Normal') { $st.Font.Size = 11; $st.ParagraphFormat.SpaceAfter = 6; $st.ParagraphFormat.LineSpacingRule = 0 }
		} catch { }
	}
	$doc.Styles.Item('Heading 1').Font.Size = 16
	$doc.Styles.Item('Heading 2').Font.Size = 13
	$doc.Styles.Item('Title').Font.Size = 20
	return $doc
}
function End-Range($doc) { $r = $doc.Content; $r.Collapse(0); return $r }
function Add-Para($doc, [string]$text, [string]$style = 'Normal') {
	$r = End-Range $doc
	$r.InsertAfter($text + "`r")
	$p = $doc.Paragraphs.Item($doc.Paragraphs.Count - 1)
	try { $p.Style = $style } catch { }
	return $p
}

# Where the first character of a paragraph is drawn. Range.Information(5) is NOT reliable for this:
# for a collapsed range, and often for a one-character range, it returns the column edge (72 pt for an
# indented or centred paragraph). Selection.Information(5) after selecting the character reports the
# caret position, which is where the text really starts (it agrees with Information(7), the offset
# from the text boundary, added to the column edge).
function Get-CaretX($word, $doc, [int]$start, [int]$end) {
	$doc.Range($start, [Math]::Min($start + 1, $end)).Select()
	return [Math]::Round([double]$word.Selection.Information(5), 1)
}

function Get-Facts($word, $doc, [string]$name) {
	$paras = @()
	$i = 0
	foreach ($p in $doc.Paragraphs) {
		$i++
		$r = $p.Range
		$text = [string]$r.Text
		$start = $r.Start; $end = $r.End
		$first = $doc.Range($start, $start)
		$entry = [ordered]@{
			index    = $i
			text     = $text.TrimEnd("`r", "`a").Substring(0, [Math]::Min(32, $text.TrimEnd("`r", "`a").Length))
			length   = $end - $start
			inTable  = [bool]$r.Information(12)
			style    = [string]$p.Style.NameLocal
			page     = [int]$first.Information(3)
			xPt      = Get-CaretX $word $doc $start $end
			yPt      = [Math]::Round([double]$first.Information(6), 1)
		}
		# Line starts: scan characters, a new line begins where the vertical position or page changes
		# (or the horizontal position steps back on the same row, as in columns).
		$starts = @(0)
		$lastY = $entry.yPt; $lastPage = $entry.page; $lastX = $entry.xPt
		$limit = $end - $start
		if (-not $entry.inTable -or $true) {
			for ($c = 1; $c -lt $limit; $c++) {
				$cr = $doc.Range($start + $c, $start + $c)
				$y = [Math]::Round([double]$cr.Information(6), 1); $pg = [int]$cr.Information(3)
				if ($y -ne $lastY -or $pg -ne $lastPage) { $starts += $c }
				$lastY = $y; $lastPage = $pg
			}
		}
		$entry.lineStarts = $starts
		$entry.lineCount = $starts.Count
		$entry.endPage = [int]$doc.Range([Math]::Max($start, $end - 1), [Math]::Max($start, $end - 1)).Information(3)
		$paras += $entry
	}
	$shapes = @()
	foreach ($s in $doc.Shapes) {
		$shapes += [ordered]@{
			name = [string]$s.Name; type = [int]$s.Type; page = [int]$s.Anchor.Information(3)
			leftPt = [Math]::Round([double]$s.Left, 1); topPt = [Math]::Round([double]$s.Top, 1)
			widthPt = [Math]::Round([double]$s.Width, 1); heightPt = [Math]::Round([double]$s.Height, 1)
			relH = [int]$s.RelativeHorizontalPosition; relV = [int]$s.RelativeVerticalPosition
			wrap = [int]$s.WrapFormat.Type
		}
	}
	$inline = @($doc.InlineShapes | ForEach-Object { [ordered]@{ type = [int]$_.Type; widthPt = [Math]::Round([double]$_.Width, 1); heightPt = [Math]::Round([double]$_.Height, 1) } })
	$sections = @($doc.Sections | ForEach-Object {
		[ordered]@{ orientation = [int]$_.PageSetup.Orientation; columns = [int]$_.PageSetup.TextColumns.Count; pageWidthPt = [double]$_.PageSetup.PageWidth; pageHeightPt = [double]$_.PageSetup.PageHeight; start = [int]$_.PageSetup.SectionStart }
	})
	return [ordered]@{
		name = $name
		word = [string]$word.Build
		font = $font
		pageCount = [int]$doc.ComputeStatistics(2)
		lineCount = [int]$doc.ComputeStatistics(1)
		footnotes = [int]$doc.Footnotes.Count
		endnotes = [int]$doc.Endnotes.Count
		tables = [int]$doc.Tables.Count
		comments = [int]$doc.Comments.Count
		revisions = [int]$doc.Revisions.Count
		sections = $sections
		paragraphs = $paras
		shapes = $shapes
		inlineShapes = $inline
	}
}

function Export-Pages($word, $doc, [string]$name) {
	$view = $doc.ActiveWindow.View
	$view.Type = 3
	$count = [int]$doc.ComputeStatistics(2)
	$files = @()
	for ($n = 1; $n -le $count; $n++) {
		$page = $doc.ActiveWindow.ActivePane.Pages.Item($n)
		$bytes = [byte[]]$page.EnhMetaFileBits
		$ms = New-Object IO.MemoryStream (, $bytes)
		$meta = New-Object Drawing.Imaging.Metafile $ms
		$w = [int][Math]::Round(612 / 72 * $Dpi); $h = [int][Math]::Round(792 / 72 * $Dpi)
		if ($meta.Width -gt $meta.Height * 1.2) { $t = $w; $w = $h; $h = $t }
		$bm = New-Object Drawing.Bitmap $w, $h
		$gr = [Drawing.Graphics]::FromImage($bm)
		$gr.Clear([Drawing.Color]::White)
		$gr.InterpolationMode = 'HighQualityBicubic'
		$gr.DrawImage($meta, 0, 0, $w, $h)
		$gr.Dispose()
		$file = Join-Path $OutDir ("{0}-p{1}.png" -f $name, $n)
		$bm.Save($file, [Drawing.Imaging.ImageFormat]::Png); $bm.Dispose(); $meta.Dispose(); $ms.Dispose()
		$files += (Split-Path -Leaf $file)
	}
	return $files
}

$builders = [ordered]@{}
$builders['01-headings-styles'] = {
	param($d)
	Add-Para $d 'Quarterly Report' 'Title' | Out-Null
	Add-Para $d 'Overview' 'Heading 1' | Out-Null
	Add-Para $d (Body 6) | Out-Null
	Add-Para $d 'Details' 'Heading 2' | Out-Null
	$p = Add-Para $d (Body 4)
	$p.Range.Font.Bold = 1; $p.Alignment = 3; $p.SpaceBefore = 12; $p.LeftIndent = 36; $p.FirstLineIndent = 18
	Add-Para $d 'Closing' 'Heading 1' | Out-Null
	$p = Add-Para $d (Body 3)
	$p.Range.Font.Italic = 1; $p.Alignment = 1
	$p = Add-Para $d (Body 12); $p.LineSpacingRule = 5; $p.LineSpacing = 18
}
$builders['02-lists-numbering'] = {
	param($d)
	Add-Para $d 'Shopping' 'Heading 1' | Out-Null
	foreach ($t in 'Apples and pears from the market stall near the old station', 'Bread', 'Milk') {
		$p = Add-Para $d $t; $p.Range.ListFormat.ApplyBulletDefault()
	}
	Add-Para $d 'Steps' 'Heading 1' | Out-Null
	$first = $null
	foreach ($t in 'Open the box and check every part against the packing list supplied inside', 'Attach the legs', 'Turn the table over') {
		$p = Add-Para $d $t
		$p.Range.ListFormat.ApplyListTemplate($d.Application.ListGalleries.Item(2).ListTemplates.Item(1), $(if ($first) { $true } else { $false }), 0)
		$first = $true
	}
	$p = Add-Para $d 'Outline level one'
	$p.Range.ListFormat.ApplyListTemplate($d.Application.ListGalleries.Item(3).ListTemplates.Item(1), $false, 0)
	$p = Add-Para $d 'Outline level two with enough words to wrap onto a second line of the page'
	$p.Range.ListFormat.ApplyListTemplate($d.Application.ListGalleries.Item(3).ListTemplates.Item(1), $true, 0)
	$p.Range.ListFormat.ListLevelNumber = 2
}
$builders['03-tables-merged'] = {
	param($d)
	Add-Para $d 'Table with merged cells' 'Heading 1' | Out-Null
	$r = End-Range $d
	$t = $d.Tables.Add($r, 4, 3)
	$t.Borders.Enable = 1
	$t.PreferredWidthType = 3; $t.PreferredWidth = 468
	for ($i = 1; $i -le 4; $i++) { for ($j = 1; $j -le 3; $j++) { $t.Cell($i, $j).Range.Text = "R$i C$j alpha beta gamma delta" } }
	$t.Rows.Item(1).HeadingFormat = -1
	$t.Cell(1, 1).Merge($t.Cell(1, 2))
	$t.Cell(2, 3).Merge($t.Cell(3, 3))
	$t.Cell(1, 1).Shading.BackgroundPatternColor = 14277081
	$d.Content.InsertParagraphAfter()
	Add-Para $d (Body 3) | Out-Null
}
$builders['04-header-footer'] = {
	param($d)
	$sec = $d.Sections.Item(1)
	$sec.Headers.Item(1).Range.Text = 'Confidential draft'
	$sec.Footers.Item(1).Range.Text = 'Page '
	$fr = $sec.Footers.Item(1).Range; $fr.Collapse(0)
	$d.Fields.Add($fr, 33) | Out-Null
	for ($i = 1; $i -le 3; $i++) {
		Add-Para $d "Section text $i" 'Heading 1' | Out-Null
		Add-Para $d (Body 14) | Out-Null
	}
}
$builders['05-footnotes-endnotes'] = {
	param($d)
	Add-Para $d 'Notes' 'Heading 1' | Out-Null
	$p = Add-Para $d (Body 6)
	$r = $p.Range; $r.SetRange($r.Start + 20, $r.Start + 20)
	$d.Footnotes.Add($r, [Type]::Missing, 'First footnote text that is long enough to wrap onto a second line.') | Out-Null
	$p = Add-Para $d (Body 5)
	$r = $p.Range; $r.SetRange($r.Start + 40, $r.Start + 40)
	$d.Footnotes.Add($r, [Type]::Missing, 'Second footnote.') | Out-Null
	$p = Add-Para $d (Body 4)
	$r = $p.Range; $r.SetRange($r.Start + 10, $r.Start + 10)
	$d.Endnotes.Add($r, [Type]::Missing, 'An endnote at the end of the document.') | Out-Null
}
$builders['06-sections-columns'] = {
	param($d)
	Add-Para $d 'Portrait first section' 'Heading 1' | Out-Null
	Add-Para $d (Body 4) | Out-Null
	$d.Content.InsertParagraphAfter()
	$r = End-Range $d; $r.InsertBreak(2)
	$s2 = $d.Sections.Item($d.Sections.Count)
	$s2.PageSetup.Orientation = 1
	Add-Para $d 'Landscape section' 'Heading 1' | Out-Null
	Add-Para $d (Body 5) | Out-Null
	$d.Content.InsertParagraphAfter()
	$r = End-Range $d; $r.InsertBreak(2)
	$s3 = $d.Sections.Item($d.Sections.Count)
	$s3.PageSetup.Orientation = 0
	$s3.PageSetup.TextColumns.SetCount(2)
	Add-Para $d 'Two columns' 'Heading 1' | Out-Null
	Add-Para $d (Body 18) | Out-Null
}
$builders['07-images'] = {
	param($d)
	Add-Para $d 'Pictures' 'Heading 1' | Out-Null
	$p = Add-Para $d 'Inline picture: '
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
	$d.InlineShapes.AddPicture($png, $false, $true, $r) | Out-Null
	$p = Add-Para $d (Body 8)
	$sh = $d.Shapes.AddPicture($png, $false, $true, 330, 200, 120, 75, $p.Range)
	$sh.RelativeHorizontalPosition = 1; $sh.RelativeVerticalPosition = 1
	$sh.Left = 330; $sh.Top = 200
	$sh.WrapFormat.Type = 0
	Add-Para $d (Body 10) | Out-Null
}
$builders['08-equation'] = {
	param($d)
	Add-Para $d 'Equations' 'Heading 1' | Out-Null
	Add-Para $d (Body 2) | Out-Null
	$p = Add-Para $d 'x=(-b+sqrt(b^2-4ac))/(2a)'
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null
	$d.OMaths.Add($r) | Out-Null; $d.OMaths.BuildUp()
	$p = Add-Para $d 'a_n=sum_(k=1)^n k^2'
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null
	$d.OMaths.Add($r) | Out-Null; $d.OMaths.BuildUp()
	Add-Para $d (Body 3) | Out-Null
}
# Display and inline equations of several shapes, each followed by a plain paragraph, so the height
# Word gives an equation paragraph shows in where the next paragraph starts.
function Add-Equation($d, [string]$linear, [string]$before = '', [string]$after = '') {
	$p = Add-Para $d ($before + $linear + $after)
	$r = $p.Range; $r.SetRange($r.Start + $before.Length, $r.Start + $before.Length + $linear.Length)
	$d.OMaths.Add($r) | Out-Null
}
$builders['13-equation-shapes'] = {
	param($d)
	Add-Para $d 'Shapes' 'Heading 1' | Out-Null
	foreach ($e in 'x=1', 'a_n=k^2', '(a+b)/(c+d)', "$([char]0x2211)_(k=1)^n k^2 ", "$([char]0x25A0)(1&2@3&4)", "$([char]0x221A)(b^2-4ac)") {
		Add-Equation $d $e
		Add-Para $d 'After the equation.' | Out-Null
	}
	Add-Equation $d '(a+b)/(c+d)' 'Inline fraction ' ' and more text.'
	Add-Para $d 'Last paragraph.' | Out-Null
	$d.OMaths.BuildUp()
}
$builders['09-fields-toc'] = {
	param($d)
	$p = Add-Para $d 'Contents' 'Heading 1'; $p.OutlineLevel = 10
	Add-Para $d 'TOCPLACEHOLDER' | Out-Null
	foreach ($t in 'Introduction', 'Method', 'Results') {
		$r = End-Range $d; $r.InsertBreak(7)
		Add-Para $d $t 'Heading 1' | Out-Null
		Add-Para $d (Body 5) | Out-Null
	}
	$ph = $d.Paragraphs.Item(2).Range; $ph.MoveEnd(1, -1) | Out-Null
	$ph.Text = ''
	$d.TablesOfContents.Add($ph, $true, 1, 1) | Out-Null
	$p = Add-Para $d 'Date field: '
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
	$d.Fields.Add($r, 31) | Out-Null
	$d.Fields.Update() | Out-Null
	$d.TablesOfContents.Item(1).Update()
}
$builders['10-tracked-comments'] = {
	param($d)
	Add-Para $d 'Review' 'Heading 1' | Out-Null
	$p = Add-Para $d (Body 5)
	$d.TrackRevisions = $true
	$r = $p.Range; $r.SetRange($r.Start + 4, $r.Start + 9); $r.Text = 'swift'
	$r = $p.Range; $r.SetRange($r.Start + 30, $r.Start + 30); $r.InsertAfter(' inserted words ')
	$r = $p.Range; $r.SetRange($r.Start + 100, $r.Start + 112); $r.Delete() | Out-Null
	$d.TrackRevisions = $false
	$r = $p.Range; $r.SetRange($r.Start + 60, $r.Start + 70)
	$d.Comments.Add($r, 'Please check this wording.') | Out-Null
	Add-Para $d (Body 3) | Out-Null
}
$builders['11-content-controls'] = {
	param($d)
	Add-Para $d 'Form' 'Heading 1' | Out-Null
	$p = Add-Para $d 'Name: '
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
	$cc = $d.ContentControls.Add(1, $r); $cc.Title = 'Name'; $cc.Tag = 'name'; 
	$p = Add-Para $d 'Colour: '
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
	$cc = $d.ContentControls.Add(4, $r); $cc.Title = 'Colour'
	$cc.DropdownListEntries.Add('Red') | Out-Null; $cc.DropdownListEntries.Add('Green') | Out-Null; $cc.DropdownListEntries.Add('Blue') | Out-Null
	$cc.DropdownListEntries.Item(2).Select()
	$p = Add-Para $d 'Agree: '
	$r = $p.Range; $r.MoveEnd(1, -1) | Out-Null; $r.Collapse(0)
	$cc = $d.ContentControls.Add(8, $r); $cc.Checked = $true
	Add-Para $d (Body 4) | Out-Null
}
$builders['12-chart'] = {
	param($d)
	Add-Para $d 'Chart' 'Heading 1' | Out-Null
	$p = Add-Para $d (Body 2)
	$r = End-Range $d
	$s = $d.InlineShapes.AddChart2(-1, 51, $r)
	try { $s.Chart.ChartData.Workbook.Close() } catch { }
	Add-Para $d (Body 3) | Out-Null
}

$word = $null
$summary = @()
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $true
	$word.DisplayAlerts = 0
	foreach ($name in $builders.Keys) {
		if ($Only.Count -gt 0 -and $Only -notcontains $name) { continue }
		$doc = $null
		try {
			$doc = Initialize-Doc $word
			& $builders[$name] $doc
			$doc.Repaginate()
			$docx = Join-Path $OutDir "$name.docx"
			$doc.SaveAs2($docx, 16)
			$facts = Get-Facts $word $doc $name
			$facts.pages = @(Export-Pages $word $doc $name)
			$facts.dpi = $Dpi
			($facts | ConvertTo-Json -Depth 8) | Set-Content -Encoding utf8 (Join-Path $OutDir "$name.facts.json")
			$summary += "{0}: {1} page(s)" -f $name, $facts.pageCount
		} catch {
			$summary += "{0}: FAILED {1} at line {2}" -f $name, $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber
			$script:failed = $true
		} finally {
			if ($doc) { try { $doc.Close(0) } catch { } }
		}
	}
} finally {
	if ($word) { try { $word.Quit(0) } catch { } }
}
$summary
if ($script:failed) { exit 1 }
