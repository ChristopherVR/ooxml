# Measures how desktop Word sizes table rows (manual; needs desktop Word on Windows, not CI).
#
#   pwsh -File src/core/scripts/docx-word-table-probe.ps1 [-Out <file.json>]
#
# Small hand-written documents (Courier New 11 pt, no default spacing) opened in Word; for every
# labelled paragraph the page and top y (points, Range.Information) are recorded. Groups:
#   vmerge   a vertical merge whose text is shorter than, equal to and taller than its rows, over
#            two and three rows, merged cell in the first or the second column, with and without
#            borders, so the row height Word gives the last spanned row shows in the next row's y
#   borders  one- and two-row tables with no border, single 0.5 / 1.5 / 3 pt and double borders
#   rowh     atLeast and exact row heights against the content height
# The result is committed as src/core/docx/__fixtures__/word-table/table-probe.json.
param(
	[string]$Out = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/word-table/table-probe.json'),
	[int]$WaitMinutes = 20
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'docx-word-probe-common.ps1')
New-Item -ItemType Directory -Force (Split-Path -Parent $Out) | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-table-work'
New-Item -ItemType Directory -Force $work | Out-Null
Wait-NoWord $WaitMinutes

function Cell([string]$paras, [string]$tcPr = '') { "<w:tc><w:tcPr><w:tcW w:w=`"3000`" w:type=`"dxa`"/>$tcPr</w:tcPr>$paras</w:tc>" }
function Row([string]$cells, [string]$trPr = '') { "<w:tr><w:trPr>$trPr</w:trPr>$cells</w:tr>" }
function Tbl([string]$rows, [int]$cols, [string]$border) {
	$b = if ($border) { "<w:tblBorders>" + (('top', 'left', 'bottom', 'right', 'insideH', 'insideV' | ForEach-Object { "<w:$_ w:val=`"$($border.Split(':')[0])`" w:sz=`"$($border.Split(':')[1])`" w:space=`"0`" w:color=`"auto`"/>" }) -join '') + "</w:tblBorders>" } else { '' }
	$grid = ('<w:gridCol w:w="3000"/>' * $cols)
	"<w:tbl><w:tblPr><w:tblW w:w=`"$(3000 * $cols)`" w:type=`"dxa`"/>$b<w:tblLayout w:type=`"fixed`"/><w:tblCellMar><w:top w:w=`"0`" w:type=`"dxa`"/><w:bottom w:w=`"0`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr><w:tblGrid>$grid</w:tblGrid>$rows</w:tbl>"
}
function Lines([string]$id, [int]$n, [string]$after = '0') { (1..$n | ForEach-Object { Para "${id}l$_" "w:after=`"$after`"" }) -join '' }

$docs = [ordered]@{}
$borders = [ordered]@{ n = ''; s = 'single:4' }
# vmerge: [mergedColumn 0|1] [rows spanned 2|3] [merged text lines] with other cells 1 line.
foreach ($bk in $borders.Keys) { foreach ($col in 0, 1) { foreach ($span in 2, 3) { foreach ($n in 1, 2, 4, 7) {
	$id = "vm${bk}_c${col}_s${span}_n${n}"
	$rows = ''
	for ($r = 1; $r -le $span + 1; $r++) {
		$cells = @()
		for ($c = 0; $c -lt 2; $c++) {
			if ($c -eq $col -and $r -le $span) {
				$content = if ($r -eq 1) { Lines "${id}m" $n } else { Para "${id}e$r" }
				$pr = if ($r -eq 1) { '<w:vMerge w:val="restart"/>' } else { '<w:vMerge/>' }
				$cells += Cell $content $pr
			} else { $cells += Cell (Para "${id}r${r}c$c") }
		}
		$rows += Row ($cells -join '')
	}
	$docs[$id] = (Para "${id}top") + (Tbl $rows 2 $borders[$bk]) + (Para "${id}after")
} } } }
# vmerge with paragraph spacing like the corpus (6 pt after everywhere), 2-row span.
foreach ($n in 1, 2) {
	$id = "vmsp_n$n"
	$rows = (Row ((Cell (Lines "${id}m" $n '120') '<w:vMerge w:val="restart"/>') + (Cell (Para "${id}r1" 'w:after="120"')))) +
		(Row ((Cell (Para "${id}e2" 'w:after="120"') '<w:vMerge/>') + (Cell (Para "${id}r2" 'w:after="120"')))) +
		(Row ((Cell (Para "${id}r3a" 'w:after="120"')) + (Cell (Para "${id}r3b" 'w:after="120"'))))
	$docs[$id] = (Para "${id}top") + (Tbl $rows 2 'single:4') + (Para "${id}after")
}
# borders: one and two rows, each border style, row text 1 line with 6 pt after.
$styles = [ordered]@{ none = ''; s4 = 'single:4'; s12 = 'single:12'; s24 = 'single:24'; dbl4 = 'double:4' }
foreach ($k in $styles.Keys) {
	$id = "bd_$k"
	$rows = (Row (Cell (Para "${id}r1" 'w:after="120"'))) + (Row (Cell (Para "${id}r2" 'w:after="120"'))) + (Row (Cell (Para "${id}r3" 'w:after="120"')))
	$docs[$id] = (Para "${id}top") + (Tbl $rows 1 $styles[$k]) + (Para "${id}after")
}
# rowh: atLeast / exact heights 40 pt and 10 pt against 1 line of text.
foreach ($rule in 'atLeast', 'exact') { foreach ($h in 800, 200) {
	$id = "rh_${rule}_$h"
	$tr = "<w:trHeight w:val=`"$h`" w:hRule=`"$rule`"/>"
	$rows = (Row (Cell (Para "${id}r1")) $tr) + (Row (Cell (Para "${id}r2")) $tr)
	$docs[$id] = (Para "${id}top") + (Tbl $rows 1 'single:4') + (Para "${id}after")
} }

$groupOf = { param($k) if ($k -like 'vm*') { 'vmerge' } elseif ($k -like 'bd_*') { 'borders' } else { 'rowh' } }
$groups = [ordered]@{ vmerge = [ordered]@{}; borders = [ordered]@{}; rowh = [ordered]@{} }
$word = $null; $build = ''
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $true; $word.DisplayAlerts = 0; $build = [string]$word.Build
	foreach ($name in $docs.Keys) {
		$path = Join-Path $work "$name.docx"
		Write-Docx $path $docs[$name] 15 $false $bodySect
		$doc = $word.Documents.Open($path, $false, $true)
		try {
			$doc.Repaginate()
			$rows = @()
			foreach ($p in $doc.Paragraphs) {
				$r = $p.Range
				$first = $doc.Range($r.Start, $r.Start)
				$label = ([string]$r.Text).TrimEnd("`r", "`a", "`f")
				if (-not $label) { continue }
				$rows += [ordered]@{ label = $label; page = [int]$first.Information(3); y = [Math]::Round([double]$first.Information(6), 2) }
			}
			$groups[(& $groupOf $name)][$name] = $rows
		} finally { $doc.Close(0) }
	}
} finally {
	if ($word) { try { $word.Quit(0) } catch { } }
}
@{ word = $build; groups = $groups } | ConvertTo-Json -Depth 8 -Compress | Set-Content -Encoding utf8 $Out
"wrote $Out"
