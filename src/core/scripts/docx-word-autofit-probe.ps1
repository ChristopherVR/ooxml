# Records how desktop Word sizes AutoFit table columns (manual; needs desktop Word on Windows, not CI).
#
#   pwsh -File src/core/scripts/docx-word-autofit-probe.ps1 [-Out <dir>]
#
# Small hand-written documents (Courier New 11 pt, cell margins 108 twips, Letter page with 1 in
# margins) whose tables carry a deliberately wrong grid (1000 twips per column). Word recomputes an
# AutoFit table's columns when it opens a document, so the grid it saves back is its own layout.
# Two `consistent-*` cases carry the grid Word would write itself (equal to the cells' widths).
# For every case the source document is written to <dir>/<case>.docx and Word's saved
# `w:tblGrid/w:gridCol` widths (twips) are recorded in <dir>/word-grids.json.
param(
	[string]$Out = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/table-autofit'),
	[int]$WaitMinutes = 20
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'docx-word-probe-common.ps1')
New-Item -ItemType Directory -Force $Out | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-autofit-work'
New-Item -ItemType Directory -Force $work | Out-Null

function C([string]$text, [string]$tcPr = '') { "<w:tc><w:tcPr>$tcPr</w:tcPr><w:p><w:r><w:t xml:space=`"preserve`">$text</w:t></w:r></w:p></w:tc>" }
function TRow([string[]]$cells) { '<w:tr>' + ($cells -join '') + '</w:tr>' }
function T([string]$tblPr, [string]$rows, [int]$cols = 3, [int]$gridWidth = 1000) {
	$grid = "<w:gridCol w:w=`"$gridWidth`"/>" * $cols
	"<w:tbl><w:tblPr>$tblPr<w:tblCellMar><w:left w:w=`"108`" w:type=`"dxa`"/><w:right w:w=`"108`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr><w:tblGrid>$grid</w:tblGrid>$rows</w:tbl>"
}
# A table without its own cell margins, so they come from the styles part.
function TBare([string]$tblPr, [string]$rows) {
	"<w:tbl><w:tblPr>$tblPr</w:tblPr><w:tblGrid>" + ('<w:gridCol w:w="1000"/>' * 3) + "</w:tblGrid>$rows</w:tbl>"
}
function TableStyle([string]$id, [string]$inner, [string]$extra = '') {
	"<w:style w:type=`"table`" $extra w:styleId=`"$id`"><w:name w:val=`"$id`"/>$inner</w:style>"
}
function Mar([int]$left, [int]$right) {
	"<w:tblPr><w:tblCellMar><w:left w:w=`"$left`" w:type=`"dxa`"/><w:right w:w=`"$right`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr>"
}
function Set-Entry([string]$path, [string]$name, [string]$text) {
	$zip = [IO.Compression.ZipFile]::Open($path, 'Update')
	try {
		$zip.GetEntry($name).Delete()
		$stream = $zip.CreateEntry($name).Open()
		$bytes = [Text.UTF8Encoding]::new($false).GetBytes($text)
		$stream.Write($bytes, 0, $bytes.Length); $stream.Close()
	} finally { $zip.Dispose() }
}
function Get-Entry([string]$path, [string]$name) {
	$zip = [IO.Compression.ZipFile]::OpenRead($path)
	try {
		$reader = New-Object IO.StreamReader(($zip.GetEntry($name)).Open())
		$text = $reader.ReadToEnd(); $reader.Close(); $text
	} finally { $zip.Dispose() }
}
$auto = '<w:tblW w:w="0" w:type="auto"/>'
$long = 'the quick brown fox jumps over the lazy dog and keeps running past the river bank until night'
$mid = 'alpha beta gamma delta epsilon zeta eta theta'
$span2 = '<w:gridSpan w:val="2"/>'
function Pref([int]$w, [string]$type = 'dxa') { "<w:tcW w:w=`"$w`" w:type=`"$type`"/>" }

$cases = [ordered]@{
	'max-fits'          = T $auto (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'interpolate'       = T $auto (TRow (C 'A'), (C 'Supercalifragilistic'), (C $long))
	'interpolate-two'   = T $auto (TRow (C 'Short'), (C $mid), (C $long))
	'interpolate-rows'  = T $auto ((TRow (C 'A'), (C 'Hello world'), (C 'xyz')) + (TRow (C 'Longer text here'), (C 'b'), (C $mid)))
	'table-pct-100'     = T '<w:tblW w:w="5000" w:type="pct"/>' (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'table-pct-50'      = T '<w:tblW w:w="2500" w:type="pct"/>' (TRow (C 'A'), (C 'Hello world'), (C $mid))
	'table-dxa-narrow'  = T '<w:tblW w:w="6000" w:type="dxa"/>' (TRow (C 'A'), (C 'Hello world'), (C $long))
	'table-dxa-wide'    = T '<w:tblW w:w="8000" w:type="dxa"/>' (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'table-dxa-tiny'    = T '<w:tblW w:w="1000" w:type="dxa"/>' (TRow (C 'Hello world'), (C 'Supercalifragilistic'), (C 'xyz'))
	'cell-dxa'          = T $auto (TRow (C 'A' (Pref 3000)), (C 'Hello world'), (C $long))
	'cell-dxa-small'    = T $auto (TRow (C 'Supercalifragilistic' (Pref 300)), (C 'Hello world'), (C $long))
	'cell-dxa-rows'     = T $auto ((TRow (C 'A' (Pref 2000)), (C 'B'), (C 'C')) + (TRow (C 'A' (Pref 3000)), (C 'B'), (C 'C')))
	'cell-dxa-all'      = T $auto (TRow (C 'A' (Pref 2000)), (C 'B' (Pref 2000)), (C $long (Pref 2000)))
	'cell-dxa-over'     = T $auto (TRow (C 'A' (Pref 5000)), (C 'B' (Pref 5000)), (C 'C' (Pref 5000)))
	'cell-dxa-unequal'  = T $auto (TRow (C 'A' (Pref 6000)), (C 'Supercalifragilisticexpialidocious' (Pref 3000)), (C 'C' (Pref 2000)))
	'cell-pct'          = T $auto (TRow (C 'A' (Pref 2000 'pct')), (C 'Hello world'), (C 'xyz'))
	'span-equal'        = T $auto ((TRow (C 'A'), (C 'B'), (C 'C')) + (TRow (C 'Antidisestablishmentarian' $span2), (C 'C')))
	'span-weighted'     = T $auto ((TRow (C 'A'), (C 'Hello world'), (C 'C')) + (TRow (C 'Antidisestablishmentarianism' $span2), (C 'C')))
	'span-text'         = T $auto ((TRow (C 'A'), (C 'Hello world'), (C 'xyz')) + (TRow (C $mid $span2), (C 'xyz')))
	'empty-cell'        = T $auto (TRow '<w:tc><w:p/></w:tc>', (C 'Hello world'), (C 'xyz'))
	'paragraph-indent'  = T $auto (TRow '<w:tc><w:p><w:pPr><w:ind w:left="720" w:right="360"/></w:pPr><w:r><w:t>Hello world</w:t></w:r></w:p></w:tc>', (C 'B'), (C 'xyz'))
	'two-paragraphs'    = T $auto (TRow '<w:tc><w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:p><w:r><w:t>Hello world again</w:t></w:r></w:p></w:tc>', (C 'B'), (C 'xyz'))
	'split-runs'        = T $auto (TRow '<w:tc><w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>wor</w:t></w:r><w:r><w:t>ld</w:t></w:r></w:p></w:tc>', (C 'B'), (C 'xyz'))
	'table-indent'      = T ($auto + '<w:tblInd w:w="1440" w:type="dxa"/>') (TRow (C 'A'), (C 'Supercalifragilistic'), (C $long))
	'fixed'             = T ($auto + '<w:tblLayout w:type="fixed"/>') (TRow (C 'A'), (C 'Supercalifragilistic'), (C $long))
	# A grid that agrees with the preferred widths, as Word itself saves it.
	'consistent-wide'   = T $auto (TRow (C 'A' (Pref 5000)), (C 'B' (Pref 5000)), (C 'C' (Pref 5000))) 3 5000
	'consistent-word'   = T $auto (TRow (C 'A' (Pref 3000)), (C 'Antidisestablishmentarianism' (Pref 3000)), (C 'C' (Pref 3000))) 3 3000
	# Cell margins from the styles part: none, the default table style, a style chain, own margins.
	'style-none'        = TBare $auto (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'style-default'     = TBare $auto (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'style-chain'       = TBare ($auto + '<w:tblStyle w:val="Child"/>') (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
	'style-own'         = TBare ($auto + '<w:tblStyle w:val="Child"/><w:tblCellMar><w:left w:w="20" w:type="dxa"/></w:tblCellMar>') (TRow (C 'A'), (C 'Hello world'), (C 'xyz'))
}
# Table styles added to the styles part (style-* cases); their Word-saved copies are kept too.
$defaultStyle = TableStyle 'Plain' (Mar 50 50) 'w:default="1"'
$chain = (TableStyle 'Base' (Mar 200 300)) + (TableStyle 'Child' '<w:basedOn w:val="Base"/>')
$tableStyles = @{ 'style-default' = $defaultStyle; 'style-chain' = $defaultStyle + $chain; 'style-own' = $defaultStyle + $chain }
# The same layouts saved in older compatibility modes (0 writes no compatibilityMode setting).
$modes = [ordered]@{}
foreach ($name in $cases.Keys) { $modes[$name] = 15 }
foreach ($name in 'interpolate', 'table-pct-100', 'span-weighted', 'cell-dxa-unequal') {
	foreach ($mode in 14, 12, 0) {
		$cases["$name-mode$mode"] = $cases[$name]
		$modes["$name-mode$mode"] = $mode
	}
}

# Another Word may be running; pass -WaitMinutes 0 to share the machine with it.
if ($WaitMinutes -gt 0) { Wait-NoWord $WaitMinutes }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$results = [ordered]@{ wordBuild = $word.Build; unit = 'twips'; cases = [ordered]@{} }
try {
	foreach ($name in $cases.Keys) {
		$source = Join-Path $Out "$name.docx"
		$saved = Join-Path $work "$name.word.docx"
		Write-Docx $source ($cases[$name] + (Para 'after')) $modes[$name] $false $bodySect
		if ($modes[$name] -eq 0) {
			# Write-Docx always writes a mode: rewrite the settings part without one.
			$zip = [IO.Compression.ZipFile]::Open($source, 'Update')
			try {
				$zip.GetEntry('word/settings.xml').Delete()
				$stream = $zip.CreateEntry('word/settings.xml').Open()
				$bytes = [Text.UTF8Encoding]::new($false).GetBytes("<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:settings $W/>")
				$stream.Write($bytes, 0, $bytes.Length); $stream.Close()
			} finally { $zip.Dispose() }
		}
		if ($tableStyles[$name]) {
			$styles = (Get-Entry $source 'word/styles.xml').Replace('</w:styles>', $tableStyles[$name] + '</w:styles>')
			Set-Entry $source 'word/styles.xml' $styles
		}
		$doc = $word.Documents.Open($source, $false, $true)
		$doc.SaveAs2($saved, 12)
		$doc.Close($false)
		if ($name -like 'style-*') { Copy-Item $saved (Join-Path $Out "$name.word.docx") -Force }
		$zip = [IO.Compression.ZipFile]::OpenRead($saved)
		try {
			$reader = New-Object IO.StreamReader(($zip.GetEntry('word/document.xml')).Open())
			$xml = $reader.ReadToEnd(); $reader.Close()
		} finally { $zip.Dispose() }
		$grid = @([regex]::Matches($xml, '<w:gridCol w:w="(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value })
		$results.cases[$name] = [ordered]@{ mode = $modes[$name]; grid = $grid }
		Write-Host "$name mode $($modes[$name]): $($grid -join ', ')"
	}
} finally { $word.Quit() }
$results | ConvertTo-Json -Depth 5 | Set-Content -Encoding utf8 (Join-Path $Out 'word-grids.json')
