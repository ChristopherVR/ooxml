# Records where desktop Word places floating tables (`w:tblpPr`) and how text wraps around them
# (manual; needs desktop Word on Windows, not CI).
#
#   pwsh -File src/core/scripts/docx-word-floating-table-probe.ps1 [-Out <dir>] [-WaitMinutes 0]
#
# Small hand-written documents (Courier New 11 pt, Letter page with 1 in margins, fixed 3000-twip
# table with 108-twip cell margins) opened in Word. For each case the source is written to
# <dir>/<case>.docx and <dir>/word-positions.json records, in points from the page's top-left
# (Range.Information): the page and position of the first table cell's text, and the start and end
# of every line of the paragraph that follows the table.
param(
	[string]$Out = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/floating-table'),
	[int]$WaitMinutes = 20
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'docx-word-probe-common.ps1')
New-Item -ItemType Directory -Force $Out | Out-Null

$words = 'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi omega'
$text = (($words + ' ') * 6).Trim()
function Floating([string]$tblpPr, [int]$rows = 2) {
	$row = '<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr>'
	"<w:tbl><w:tblPr><w:tblpPr $tblpPr/><w:tblW w:w=`"3000`" w:type=`"dxa`"/><w:tblBorders><w:top w:val=`"single`" w:sz=`"4`" w:space=`"0`" w:color=`"auto`"/><w:left w:val=`"single`" w:sz=`"4`" w:space=`"0`" w:color=`"auto`"/><w:bottom w:val=`"single`" w:sz=`"4`" w:space=`"0`" w:color=`"auto`"/><w:right w:val=`"single`" w:sz=`"4`" w:space=`"0`" w:color=`"auto`"/></w:tblBorders><w:tblLayout w:type=`"fixed`"/><w:tblCellMar><w:left w:w=`"108`" w:type=`"dxa`"/><w:right w:w=`"108`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w=`"3000`"/></w:tblGrid>" + ($row * $rows) + '</w:tbl>'
}
$dist = 'w:leftFromText="180" w:rightFromText="180" w:topFromText="0" w:bottomFromText="0"'
$cases = [ordered]@{
	'margin-left'   = (Para 'Before') + (Floating "$dist w:vertAnchor=`"text`" w:horzAnchor=`"margin`" w:tblpX=`"0`" w:tblpY=`"0`"") + (Para $text)
	'margin-middle' = (Para 'Before') + (Floating "$dist w:vertAnchor=`"text`" w:horzAnchor=`"margin`" w:tblpX=`"3000`" w:tblpY=`"0`"") + (Para $text)
	'page-right'    = (Para 'Before') + (Floating "$dist w:vertAnchor=`"page`" w:horzAnchor=`"page`" w:tblpXSpec=`"right`" w:tblpY=`"2880`"") + (Para $text)
	'margin-center' = (Para 'Before') + (Floating "$dist w:vertAnchor=`"margin`" w:horzAnchor=`"margin`" w:tblpXSpec=`"center`" w:tblpYSpec=`"top`"") + (Para $text)
	'text-offset'   = (Para 'One') + (Para 'Two') + (Floating "$dist w:vertAnchor=`"text`" w:horzAnchor=`"text`" w:tblpX=`"720`" w:tblpY=`"360`"") + (Para $text)
	'defaults'      = (Para 'Before') + (Floating 'w:tblpX="1440" w:tblpY="720"') + (Para $text)
	'distances'     = (Para 'Before') + (Floating 'w:leftFromText="0" w:rightFromText="720" w:topFromText="360" w:bottomFromText="360" w:vertAnchor="text" w:horzAnchor="margin" w:tblpX="0" w:tblpY="0"') + (Para $text)
}

if ($WaitMinutes -gt 0) { Wait-NoWord $WaitMinutes }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$results = [ordered]@{ wordBuild = $word.Build; unit = 'points from the page top-left'; cases = [ordered]@{} }
function Pos($range) { @([math]::Round($range.Information(5), 2), [math]::Round($range.Information(6), 2), $range.Information(3)) }
try {
	foreach ($name in $cases.Keys) {
		$source = Join-Path $Out "$name.docx"
		Write-Docx $source $cases[$name] 15 $false $bodySect
		$doc = $word.Documents.Open($source, $false, $true)
		$doc.ActiveWindow.View.Type = 3
		$cell = $doc.Tables.Item(1).Cell(1, 1).Range
		$cell.Collapse(1)
		$paragraph = $doc.Paragraphs.Item($doc.Paragraphs.Count)
		$lines = @()
		$current = $null
		foreach ($wordRange in $paragraph.Range.Words) {
			$start = $doc.Range($wordRange.Start, $wordRange.Start)
			$p = Pos $start
			$text = $wordRange.Text.TrimEnd()
			$endX = [math]::Round($start.Information(5) + 6.6 * $text.Length, 2)
			if (-not $current -or $current.y -ne $p[1] -or $current.page -ne $p[2]) {
				if ($current) { $lines += $current }
				$current = [ordered]@{ page = $p[2]; y = $p[1]; x = $p[0]; endX = $endX }
			} else { $current.endX = $endX }
		}
		if ($current) { $lines += $current }
		$c = Pos $cell
		$results.cases[$name] = [ordered]@{ cellText = [ordered]@{ x = $c[0]; y = $c[1]; page = $c[2] }; lines = $lines }
		Write-Host "$name cell=$($c -join ',') lines=$($lines.Count) first=$($lines[0].x),$($lines[0].y)"
		$doc.Close($false)
	}
} finally { $word.Quit() }
$results | ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $Out 'word-positions.json')
