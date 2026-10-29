# Manual probe (Windows, desktop Word, not run in CI): applies table structure edits in real Word
# on small tables and records the resulting grid, so the docx table-structure engine can match it.
#
#   pwsh -File table-structure-word-probe.ps1 -Out <word-table-structure.json> [-Dir <scratch>]
#
# Each scenario builds a table, performs one edit through Word's object model, saves a .docx and
# reads w:tbl back as a compact grid: one string per cell, `text[:span][^|v]{tcW}`, where `^` is
# vMerge restart and `|` is vMerge continue.
param(
	[Parameter(Mandatory)][string]$Out,
	[string]$Dir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-table-probe')
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Dir | Out-Null
$deadline = (Get-Date).AddMinutes(15)
while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
	if ((Get-Date) -gt $deadline) { throw 'Word stays busy' }
	Start-Sleep -Seconds 30
}
$w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

function Read-Grid([string]$path, [int]$index = 0) {
	Add-Type -AssemblyName System.IO.Compression.FileSystem
	# Word keeps the file open for writing while the document is open: read it with shared access.
	$stream = New-Object IO.FileStream($path, 'Open', 'Read', 'ReadWrite')
	$zip = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read)
	try {
		$entry = $zip.GetEntry('word/document.xml')
		$reader = New-Object IO.StreamReader($entry.Open())
		$xml = [xml]$reader.ReadToEnd()
		$reader.Close()
	} finally { $zip.Dispose(); $stream.Dispose() }
	$ns = New-Object Xml.XmlNamespaceManager($xml.NameTable)
	$ns.AddNamespace('w', $w)
	$tbls = $xml.SelectNodes('//w:tbl[not(ancestor::w:tbl)]', $ns)
	return (Table-Of $tbls[$index] $ns)
}
function Table-Of($tbl, $ns) {
	$grid = @($tbl.SelectNodes('w:tblGrid/w:gridCol', $ns) | ForEach-Object { [int]$_.GetAttribute('w', $w) })
	$rows = @()
	foreach ($tr in $tbl.SelectNodes('w:tr', $ns)) {
		$cells = @()
		foreach ($tc in $tr.SelectNodes('w:tc', $ns)) {
			# Text of this cell's own paragraphs, one entry per paragraph (not of a nested table's).
			$paras = @($tc.SelectNodes('w:p', $ns) | ForEach-Object { ($_.SelectNodes('.//w:t', $ns) | ForEach-Object { $_.InnerText }) -join '' })
			$text = ($paras | Where-Object { $_ -ne '' }) -join '/'
			$s = $tc.SelectSingleNode('w:tcPr/w:gridSpan', $ns)
			$v = $tc.SelectSingleNode('w:tcPr/w:vMerge', $ns)
			$tw = $tc.SelectSingleNode('w:tcPr/w:tcW', $ns)
			$cell = $text
			if ($s) { $cell += ':' + $s.GetAttribute('val', $w) }
			if ($v) { if ($v.GetAttribute('val', $w) -eq 'restart') { $cell += '^' } else { $cell += '|' } }
			if ($tw) { $cell += '{' + $tw.GetAttribute('w', $w) + '}' }
			$nested = $tc.SelectNodes('w:tbl', $ns)
			if ($nested.Count) { $cell += '<' + ((Table-Of $nested[0] $ns | ConvertTo-Json -Compress -Depth 8)) + '>' }
			$cells += $cell
		}
		$trp = @($tr.SelectNodes('w:trPr/*', $ns) | ForEach-Object { $_.LocalName })
		$rows += , @{ cells = $cells; trPr = $trp }
	}
	return @{ grid = $grid; rows = $rows }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$results = @()
function Run-Scenario([string]$name, [string]$op, [int]$rows, [int]$cols, [scriptblock]$setup, [scriptblock]$edit) {
	$doc = $word.Documents.Add()
	$result = @{ name = $name; op = $op }
	try {
		$t = $doc.Tables.Add($doc.Range(0, 0), $rows, $cols)
		for ($r = 1; $r -le $rows; $r++) { for ($c = 1; $c -le $cols; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
		if ($setup) { & $setup $doc $t }
		$before = Join-Path $Dir "$name-before.docx"
		$doc.SaveAs2([string]$before, 16)
		$result.before = Read-Grid $before
		try {
			& $edit $doc $word.Selection
			$after = Join-Path $Dir "$name-after.docx"
			$doc.SaveAs2([string]$after, 16)
			$result.after = Read-Grid $after
		} catch { $result.error = $_.Exception.Message }
	} catch { $result.error = 'setup: ' + $_.Exception.Message }
	finally { $doc.Close(0) }
	return $result
}
function Cell($doc, $r, $c) { $doc.Tables(1).Cell($r, $c) }
$vm = { param($d, $t) $t.Cell(1, 1).Merge($t.Cell(3, 1)) }
$vm2 = { param($d, $t) $t.Cell(2, 1).Merge($t.Cell(3, 1)) }
$hm = { param($d, $t) $t.Cell(1, 1).Merge($t.Cell(1, 2)) }
try {
	$results += Run-Scenario 'row-below-plain' 'insert row below r2' 3 3 $null { param($d, $s) (Cell $d 2 1).Select(); $s.InsertRowsBelow(1) }
	$results += Run-Scenario 'row-above-plain' 'insert row above r2' 3 3 $null { param($d, $s) (Cell $d 2 1).Select(); $s.InsertRowsAbove(1) }
	$results += Run-Scenario 'row-below-in-vmerge' 'insert row below r1 (merge r1-r3 in column 1)' 3 3 $vm { param($d, $s) (Cell $d 1 2).Select(); $s.InsertRowsBelow(1) }
	$results += Run-Scenario 'row-above-in-vmerge' 'insert row above r2 (merge r1-r3 in column 1)' 3 3 $vm { param($d, $s) (Cell $d 2 2).Select(); $s.InsertRowsAbove(1) }
	$results += Run-Scenario 'row-above-vmerge-top' 'insert row above r1 (merge r1-r3 in column 1)' 3 3 $vm { param($d, $s) (Cell $d 1 2).Select(); $s.InsertRowsAbove(1) }
	$results += Run-Scenario 'row-below-vmerge-bottom' 'insert row below r3 (merge r1-r3 in column 1)' 3 3 $vm { param($d, $s) (Cell $d 3 2).Select(); $s.InsertRowsBelow(1) }
	$results += Run-Scenario 'row-below-hmerge' 'insert row below r1 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 1 1).Select(); $s.InsertRowsBelow(1) }
	$results += Run-Scenario 'row-above-hmerge' 'insert row above r1 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 1 1).Select(); $s.InsertRowsAbove(1) }
	$results += Run-Scenario 'delete-row-vmerge-top' 'delete r1 (merge r1-r3 col 1)' 3 3 $vm { param($d, $s) (Cell $d 1 2).Select(); $s.Rows.Delete() }
	$results += Run-Scenario 'delete-row-vmerge-middle' 'delete r2 (merge r1-r3 col 1)' 3 3 $vm { param($d, $s) (Cell $d 2 2).Select(); $s.Rows.Delete() }
	$results += Run-Scenario 'delete-row-vmerge-pair' 'delete r3 (merge r2-r3 col 1)' 3 3 $vm2 { param($d, $s) (Cell $d 3 2).Select(); $s.Rows.Delete() }
	$results += Run-Scenario 'delete-row-vmerge-pair-top' 'delete r2 (merge r2-r3 col 1)' 3 3 $vm2 { param($d, $s) (Cell $d 2 2).Select(); $s.Rows.Delete() }
	$results += Run-Scenario 'delete-row-hmerge' 'delete r1 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 1 1).Select(); $s.Rows.Delete() }
	$results += Run-Scenario 'col-right-plain' 'insert column right of c2' 3 3 $null { param($d, $s) (Cell $d 1 2).Select(); $s.InsertColumnsRight() }
	$results += Run-Scenario 'col-left-plain' 'insert column left of c2' 3 3 $null { param($d, $s) (Cell $d 1 2).Select(); $s.InsertColumns() }
	$results += Run-Scenario 'col-right-through-hmerge' 'insert column right of c1 from r2 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 2 1).Select(); $s.InsertColumnsRight() }
	$results += Run-Scenario 'col-left-through-hmerge' 'insert column left of c2 from r2 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 2 2).Select(); $s.InsertColumns() }
	$results += Run-Scenario 'col-right-from-hmerge' 'insert column right with the merged cell r1c1-c2 selected' 3 3 $hm { param($d, $s) (Cell $d 1 1).Select(); $s.InsertColumnsRight() }
	$results += Run-Scenario 'col-left-from-hmerge' 'insert column left with the merged cell r1c1-c2 selected' 3 3 $hm { param($d, $s) (Cell $d 1 1).Select(); $s.InsertColumns() }
	$results += Run-Scenario 'col-right-vmerge' 'insert column right of c1 (merge r1-r3 col 1)' 3 3 $vm { param($d, $s) (Cell $d 1 1).Select(); $s.InsertColumnsRight() }
	$results += Run-Scenario 'col-delete-plain' 'delete c2' 3 3 $null { param($d, $s) (Cell $d 1 2).Select(); $s.Columns.Delete() }
	$results += Run-Scenario 'col-delete-through-hmerge' 'delete c2 from r2 (r1c1-c2 merged)' 3 3 $hm { param($d, $s) (Cell $d 2 2).Select(); $s.Columns.Delete() }
	$results += Run-Scenario 'col-delete-vmerge' 'delete c1 (merge r1-r3 col 1)' 3 3 $vm { param($d, $s) (Cell $d 1 2).Select(); $s.Columns.Delete() }
	$results += Run-Scenario 'merge-2x2' 'merge r1c1..r2c2' 3 3 $null { param($d, $s) (Cell $d 1 1).Merge((Cell $d 2 2)) }
	$results += Run-Scenario 'merge-row' 'merge r1c1..r1c3' 3 3 $null { param($d, $s) (Cell $d 1 1).Merge((Cell $d 1 3)) }
	$results += Run-Scenario 'merge-column' 'merge r1c1..r3c1' 3 3 $null { param($d, $s) (Cell $d 1 1).Merge((Cell $d 3 1)) }
	$results += Run-Scenario 'merge-with-empty' 'merge r1c1..r1c2 where r1c2 is empty' 3 3 { param($d, $t) $t.Cell(1, 2).Range.Text = '' } { param($d, $s) (Cell $d 1 1).Merge((Cell $d 1 2)) }
	$results += Run-Scenario 'merge-into-vmerge' 'merge r1c2..r3c2 next to existing vmerge col 1' 3 3 $vm { param($d, $s) (Cell $d 1 2).Merge((Cell $d 3 2)) }
	$results += Run-Scenario 'split-cols' 'split r2c2 into 2 columns' 3 3 $null { param($d, $s) (Cell $d 2 2).Split(1, 2) }
	$results += Run-Scenario 'split-cols-3' 'split r2c2 into 3 columns' 3 3 $null { param($d, $s) (Cell $d 2 2).Split(1, 3) }
	$results += Run-Scenario 'split-rows' 'split r2c2 into 2 rows' 3 3 $null { param($d, $s) (Cell $d 2 2).Split(2, 1) }
	$results += Run-Scenario 'split-2x2' 'split r2c2 into 2 rows and 2 columns' 3 3 $null { param($d, $s) (Cell $d 2 2).Split(2, 2) }
	$results += Run-Scenario 'split-hmerge' 'split merged r1c1-c2 into 2 columns' 3 3 $hm { param($d, $s) (Cell $d 1 1).Split(1, 2) }
	$results += Run-Scenario 'split-hmerge-3' 'split merged r1c1-c2 into 3 columns' 3 3 $hm { param($d, $s) (Cell $d 1 1).Split(1, 3) }
	$results += Run-Scenario 'split-vmerge' 'split merged r1-r3 c1 into 3 rows' 3 3 $vm { param($d, $s) (Cell $d 1 1).Split(3, 1) }
	$results += Run-Scenario 'split-vmerge-2rows' 'split merged r1-r3 c1 into 2 rows' 3 3 $vm { param($d, $s) (Cell $d 1 1).Split(2, 1) }
	$nest = { param($d, $t) $t.Cell(1, 1).Range.Text = ''; $inner = $d.Tables.Add($t.Cell(1, 1).Range, 2, 2); for ($r = 1; $r -le 2; $r++) { for ($c = 1; $c -le 2; $c++) { $inner.Cell($r, $c).Range.Text = "n$r" + "$c" } } }
	$results += Run-Scenario 'nested-row' 'insert row below in the nested table' 2 2 $nest { param($d, $s) $d.Tables(1).Tables(1).Cell(1, 1).Select(); $s.InsertRowsBelow(1) }
	$results += Run-Scenario 'nested-col' 'insert column right in the nested table' 2 2 $nest { param($d, $s) $d.Tables(1).Tables(1).Cell(1, 1).Select(); $s.InsertColumnsRight() }
	$results += Run-Scenario 'nested-merge' 'merge nested n11..n12' 2 2 $nest { param($d, $s) $d.Tables(1).Tables(1).Cell(1, 1).Merge($d.Tables(1).Tables(1).Cell(1, 2)) }
	$results += Run-Scenario 'header-row-below' 'insert row below a repeated header row' 3 3 { param($d, $t) $t.Rows(1).HeadingFormat = -1; $t.Rows(1).Height = 40 } { param($d, $s) (Cell $d 1 1).Select(); $s.InsertRowsBelow(1) }
} finally {
	$word.Quit(0)
}
$results | ConvertTo-Json -Depth 12 | Set-Content -Encoding UTF8 $Out
