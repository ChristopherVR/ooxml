# Measures how desktop Word spaces paragraphs (manual; needs desktop Word on Windows, not run in CI).
#
#   pwsh -File src/core/scripts/docx-word-spacing-probe.ps1 [-Out <file.json>]
#
# For each compatibility mode (15, 14, 12) with and without w:doNotUseHTMLParagraphAutoSpacing,
# the script writes small documents by hand (Courier New 11 pt, no default spacing), opens them in
# Word and records the page and top y (points, Range.Information) of every paragraph:
#   grid     space after {0,6,12,18} pt against space before {0,6,12,18} pt
#   auto     beforeAutospacing / afterAutospacing against plain spacing, and the first paragraph
#   table    a paragraph before and after a table, spacing on the paragraph and the cell paragraph
#   secnext  section break (next page) between the pair, spacing {0,12,18}
#   seccont  continuous section break between the pair
#   pgbrk    a page break at the end of the first paragraph
#   list     three items each of a Symbol bullet, a Courier New bullet and a decimal list
#   pgbefore a paragraph with pageBreakBefore and space before
#   cols     the second column of a two-column section
#   natural  a paragraph with space before that lands at the top of a page by overflow
# The result is committed as src/core/docx/__fixtures__/word-spacing/spacing-probe.json.
# Always quits Word in a finally block and waits (up to 15 minutes) for a running WINWORD.EXE.
param(
	[string]$Out = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/word-spacing/spacing-probe.json'),
	[int]$WaitMinutes = 15
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force (Split-Path -Parent $Out) | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-spacing-work'
New-Item -ItemType Directory -Force $work | Out-Null
for ($waited = 0; (Get-Process WINWORD -ErrorAction SilentlyContinue); $waited += 30) {
	if ($waited -ge $WaitMinutes * 60) { throw 'WINWORD.EXE kept running; giving up.' }
	Start-Sleep -Seconds 30
}

$W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
function Tw([double]$pt) { [int]($pt * 20) }
function Para([string]$label, [string]$spacing = '', [string]$extra = '', [string]$run = '') {
	$sp = if ($spacing) { "<w:spacing $spacing/>" } else { '' }
	"<w:p><w:pPr>$sp$extra</w:pPr><w:r><w:t>$label</w:t></w:r>$run</w:p>"
}
function Sect([string]$type = 'nextPage') {
	"<w:sectPr><w:type w:val=`"$type`"/><w:pgSz w:w=`"12240`" w:h=`"15840`"/><w:pgMar w:top=`"1440`" w:right=`"1440`" w:bottom=`"1440`" w:left=`"1440`" w:header=`"720`" w:footer=`"720`" w:gutter=`"0`"/></w:sectPr>"
}
$bodySect = Sect 'nextPage'
function Write-Docx([string]$path, [string]$body, [int]$mode, [bool]$noHtml, [string]$sect) {
	$flag = if ($noHtml) { '<w:doNotUseHTMLParagraphAutoSpacing/>' } else { '' }
	$files = [ordered]@{
		'[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'
		'_rels/.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
		'word/_rels/document.xml.rels' = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
		'word/styles.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:styles $W><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=`"Courier New`" w:hAnsi=`"Courier New`" w:cs=`"Courier New`"/><w:sz w:val=`"22`"/><w:szCs w:val=`"22`"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after=`"0`" w:line=`"240`" w:lineRule=`"auto`"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type=`"paragraph`" w:default=`"1`" w:styleId=`"Normal`"><w:name w:val=`"Normal`"/></w:style></w:styles>"
		'word/numbering.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:numbering $W><w:abstractNum w:abstractNumId=`"1`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"&#xF0B7;`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr><w:rPr><w:rFonts w:ascii=`"Symbol`" w:hAnsi=`"Symbol`" w:hint=`"default`"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId=`"2`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"bullet`"/><w:lvlText w:val=`"o`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr><w:rPr><w:rFonts w:ascii=`"Courier New`" w:hAnsi=`"Courier New`"/></w:rPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId=`"3`"><w:lvl w:ilvl=`"0`"><w:start w:val=`"1`"/><w:numFmt w:val=`"decimal`"/><w:lvlText w:val=`"%1.`"/><w:lvlJc w:val=`"left`"/><w:pPr><w:ind w:left=`"720`" w:hanging=`"360`"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId=`"1`"><w:abstractNumId w:val=`"1`"/></w:num><w:num w:numId=`"2`"><w:abstractNumId w:val=`"2`"/></w:num><w:num w:numId=`"3`"><w:abstractNumId w:val=`"3`"/></w:num></w:numbering>"
		'word/settings.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:settings $W><w:compat>$flag<w:compatSetting w:name=`"compatibilityMode`" w:uri=`"http://schemas.microsoft.com/office/word`" w:val=`"$mode`"/></w:compat></w:settings>"
		'word/document.xml' = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><w:document $W><w:body>$body$sect</w:body></w:document>"
	}
	if (Test-Path $path) { Remove-Item $path }
	$zip = [IO.Compression.ZipFile]::Open($path, 'Create')
	try {
		foreach ($name in $files.Keys) {
			$entry = $zip.CreateEntry($name)
			$stream = $entry.Open()
			$bytes = [Text.UTF8Encoding]::new($false).GetBytes($files[$name])
			$stream.Write($bytes, 0, $bytes.Length); $stream.Close()
		}
	} finally { $zip.Dispose() }
}

$levels = 0, 6, 12, 18
$docs = [ordered]@{}
# grid: after x before; the first paragraph of each later pair has no space before.
$b = ''
foreach ($a in $levels) { foreach ($c in $levels) {
	$b += Para "a${a}b${c}x" "w:after=`"$(Tw $a)`""
	$b += Para "a${a}b${c}y" "w:before=`"$(Tw $c)`""
} }
$docs.grid = $b
# auto: autospacing (HTML) values against plain spacing; the first paragraph has space before 18.
$b = Para 'first18' "w:before=`"$(Tw 18)`""
$b += Para 'sep0' ''
$auto = [ordered]@{
	'autoA_autoB' = @('w:after="100" w:afterAutospacing="1"', 'w:before="100" w:beforeAutospacing="1"')
	'autoA_b0'    = @('w:after="100" w:afterAutospacing="1"', '')
	'a0_autoB'    = @('', 'w:before="100" w:beforeAutospacing="1"')
	'autoA_b6'    = @('w:after="100" w:afterAutospacing="1"', "w:before=`"$(Tw 6)`"")
	'autoA_b18'   = @('w:after="100" w:afterAutospacing="1"', "w:before=`"$(Tw 18)`"")
	'a6_autoB'    = @("w:after=`"$(Tw 6)`"", 'w:before="100" w:beforeAutospacing="1"')
	'a18_autoB'   = @("w:after=`"$(Tw 18)`"", 'w:before="100" w:beforeAutospacing="1"')
	'a0_b0'       = @('', '')
}
foreach ($k in $auto.Keys) {
	$b += Para "${k}x" $auto[$k][0]
	$b += Para "${k}y" $auto[$k][1]
}
$docs.auto = $b
# table: paragraph, one-cell table, paragraph.
$b = ''
foreach ($pa in 0, 12) { foreach ($cb in 0, 12) { foreach ($ca in 0, 12) { foreach ($pb in 0, 12) {
	$id = "pa${pa}_cb${cb}_ca${ca}_pb${pb}"
	$b += Para "${id}p" "w:after=`"$(Tw $pa)`""
	$cell = Para "${id}c" "w:before=`"$(Tw $cb)`" w:after=`"$(Tw $ca)`""
	$b += "<w:tbl><w:tblPr><w:tblW w:w=`"0`" w:type=`"auto`"/><w:tblCellMar><w:top w:w=`"0`" w:type=`"dxa`"/><w:bottom w:w=`"0`" w:type=`"dxa`"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w=`"4000`"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w=`"4000`" w:type=`"dxa`"/></w:tcPr>$cell</w:tc></w:tr></w:tbl>"
	$b += Para "${id}q" "w:before=`"$(Tw $pb)`""
} } } }
$docs.table = $b
# section breaks and page break: the pair, spacing after {0,12,18} x before {0,12,18}.
$three = 0, 12, 18
foreach ($kind in 'secnext', 'seccont') {
	$type = if ($kind -eq 'secnext') { 'nextPage' } else { 'continuous' }
	$b = ''
	foreach ($a in $three) { foreach ($c in $three) {
		$b += Para "a${a}b${c}x" "w:after=`"$(Tw $a)`"" (Sect $type)
		$b += Para "a${a}b${c}y" "w:before=`"$(Tw $c)`"" (Sect $type)
	} }
	$docs[$kind] = $b
}
$b = ''
foreach ($a in $three) { foreach ($c in $three) {
	$b += Para "a${a}b${c}x" "w:after=`"$(Tw $a)`"" '' '<w:r><w:br w:type="page"/></w:r>'
	$b += Para "a${a}b${c}y" "w:before=`"$(Tw $c)`""
} }
$docs.pgbrk = $b
# natural: 54 exact 12 pt lines fill the page; the target with space before overflows to page 2.
$b = Para 'first' 'w:line="240" w:lineRule="exact"'
for ($i = 1; $i -lt 54; $i++) { $b += Para "f$i" 'w:line="240" w:lineRule="exact"' }
$b += Para 'target18' "w:before=`"$(Tw 18)`""
$b += Para 'target0' ''
$docs.natural = $b
# list: three items of each list type (Symbol bullet, Courier New bullet, decimal), no spacing.
$b = ''
foreach ($n in 1, 2, 3) { 1..3 | ForEach-Object { $b += Para "list${n}i$_" '' "<w:numPr><w:ilvl w:val=`"0`"/><w:numId w:val=`"$n`"/></w:numPr>" } }
$docs.list = $b
# pgbefore: a paragraph with pageBreakBefore after a paragraph with space after {0,12,18}.
$b = ''
foreach ($a in $three) { foreach ($c in $three) {
	$b += Para "a${a}b${c}x" "w:after=`"$(Tw $a)`""
	$b += "<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:before=`"$(Tw $c)`"/></w:pPr><w:r><w:t>a${a}b${c}y</w:t></w:r></w:p>"
} }
$docs.pgbefore = $b
# cols: two columns; 54 exact 12 pt lines fill the first, the target with space before starts the second.
$colsSect = "<w:sectPr><w:pgSz w:w=`"12240`" w:h=`"15840`"/><w:pgMar w:top=`"1440`" w:right=`"1440`" w:bottom=`"1440`" w:left=`"1440`" w:header=`"720`" w:footer=`"720`" w:gutter=`"0`"/><w:cols w:num=`"2`" w:space=`"720`"/></w:sectPr>"
$b = ''
for ($i = 1; $i -le 54; $i++) { $b += Para "f$i" 'w:line="240" w:lineRule="exact"' }
$b += Para 'target18' "w:before=`"$(Tw 18)`""
$b += Para 'target0' ''
$docs.cols = $b

$results = @()
$word = $null
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $true; $word.DisplayAlerts = 0; $build = [string]$word.Build
	foreach ($mode in 15, 14, 12) { foreach ($noHtml in $false, $true) {
		$config = [ordered]@{ mode = $mode; noHtmlAutoSpacing = $noHtml; groups = [ordered]@{} }
		foreach ($name in $docs.Keys) {
			$path = Join-Path $work "$name-$mode-$([int]$noHtml).docx"
			Write-Docx $path $docs[$name] $mode $noHtml $(if ($name -eq 'cols') { $colsSect } else { $bodySect })
			$doc = $word.Documents.Open($path, $false, $true)
			try {
				$doc.Repaginate()
				$rows = @()
				foreach ($p in $doc.Paragraphs) {
					$r = $p.Range
					$first = $doc.Range($r.Start, $r.Start)
					$label = ([string]$r.Text).TrimEnd("`r", "`a", "`f")
					if (-not $label) { continue }
					$rows += [ordered]@{
						label = $label.Substring(0, [Math]::Min(40, $label.Length))
						page = [int]$first.Information(3)
						y = [Math]::Round([double]$first.Information(6), 2)
					}
				}
				$config.groups[$name] = $rows
				$config.compatMode = [int]$doc.CompatibilityMode
			} finally { $doc.Close(0) }
		}
		$results += $config
	} }
} finally {
	if ($word) { try { $word.Quit(0) } catch { } }
}
@{ word = $build; configs = $results } | ConvertTo-Json -Depth 8 -Compress | Set-Content -Encoding utf8 $Out
"wrote $Out"
