# Measures the line height Word gives lines holding content controls and East Asian symbol fonts
# (manual; needs desktop Word on Windows, not CI).
#
#   pwsh -File src/core/scripts/docx-word-content-control-probe.ps1 [-Out <file.json>]
#
# One document of Courier New 11 pt paragraphs with no spacing: a label paragraph (L<case>), then
# the case paragraph, so the distance between a case paragraph's top and the next label's top
# (Range.Information) is the case's line height. Cases:
#   plain      ordinary Courier text
#   gothic*    MS Gothic runs (the symbol, latin text with and without the East Asian hint)
#   cb*        check-box content controls, checked and unchecked, with MS Gothic and Courier runs
#   text/drop  plain text and drop-down list content controls with ordinary runs
#   sym*       the same symbol in Courier New and in Segoe UI Symbol
# The result is committed as src/core/docx/__fixtures__/word-content-control/line-probe.json.
param(
	[string]$Out = (Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) 'docx/__fixtures__/word-content-control/line-probe.json'),
	[int]$WaitMinutes = 20
)
$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) 'docx-word-probe-common.ps1')
New-Item -ItemType Directory -Force (Split-Path -Parent $Out) | Out-Null
$work = Join-Path ([IO.Path]::GetTempPath()) 'docx-word-content-control-work'
New-Item -ItemType Directory -Force $work | Out-Null
Wait-NoWord $WaitMinutes

function Run([string]$text, [string]$rPr = '') {
	$pr = if ($rPr) { "<w:rPr>$rPr</w:rPr>" } else { '' }
	"<w:r>$pr<w:t xml:space=`"preserve`">$text</w:t></w:r>"
}
$gothic = '<w:rFonts w:ascii="MS Gothic" w:eastAsia="MS Gothic" w:hAnsi="MS Gothic" w:hint="eastAsia"/>'
$gothicLatin = '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>'
$segoe = '<w:rFonts w:ascii="Segoe UI Symbol" w:hAnsi="Segoe UI Symbol"/>'
$box = [string][char]0x2612
$empty = [string][char]0x2610
function Check([string]$state, [string]$glyph, [string]$rPr) {
	"<w:sdt><w:sdtPr><w:id w:val=`"1001`"/><w14:checkbox xmlns:w14=`"http://schemas.microsoft.com/office/word/2010/wordml`"><w14:checked w14:val=`"$state`"/><w14:checkedState w14:val=`"2612`" w14:font=`"MS Gothic`"/><w14:uncheckedState w14:val=`"2610`" w14:font=`"MS Gothic`"/></w14:checkbox></w:sdtPr><w:sdtContent>$(Run $glyph $rPr)</w:sdtContent></w:sdt>"
}
$cases = [ordered]@{
	plain       = Run 'Plain: ok'
	gothicsym   = (Run 'Gothic: ') + (Run $box $gothic)
	gothictxt   = (Run 'Gothic: ') + (Run 'abc' $gothic)
	gothiclat   = (Run 'Gothic: ') + (Run 'abc' $gothicLatin)
	gothiconly  = Run $box $gothic
	cbchecked   = (Run 'Agree: ') + (Check '1' $box $gothic)
	cbunchecked = (Run 'Agree: ') + (Check '0' $empty $gothic)
	cbcourier   = (Run 'Agree: ') + (Check '1' $box '')
	text        = (Run 'Name: ') + '<w:sdt><w:sdtPr><w:id w:val="1002"/><w:text/></w:sdtPr><w:sdtContent>' + (Run 'Ada') + '</w:sdtContent></w:sdt>'
	dropdown    = (Run 'Colour: ') + '<w:sdt><w:sdtPr><w:id w:val="1003"/><w:dropDownList><w:listItem w:displayText="Red" w:value="Red"/></w:dropDownList></w:sdtPr><w:sdtContent>' + (Run 'Red') + '</w:sdtContent></w:sdt>'
	symcourier  = (Run 'Sym: ') + (Run $box)
	symsegoe    = (Run 'Sym: ') + (Run $box $segoe)
}
$body = ''
foreach ($k in $cases.Keys) { $body += "<w:p><w:r><w:t>L$k</w:t></w:r></w:p><w:p>$($cases[$k])</w:p>" }
$body += '<w:p><w:r><w:t>Lend</w:t></w:r></w:p>'

$rows = @()
$build = ''
$word = $null
try {
	$word = New-Object -ComObject Word.Application
	$word.Visible = $true; $word.DisplayAlerts = 0; $build = [string]$word.Build
	$path = Join-Path $work 'cc.docx'
	Write-Docx $path $body 15 $false $bodySect
	$doc = $word.Documents.Open($path, $false, $true)
	try {
		$doc.Repaginate()
		foreach ($p in $doc.Paragraphs) {
			$r = $p.Range
			$first = $doc.Range($r.Start, $r.Start)
			$label = ([string]$r.Text).TrimEnd("`r", "`a", "`f")
			$rows += [ordered]@{ label = $label; page = [int]$first.Information(3); y = [Math]::Round([double]$first.Information(6), 2) }
		}
	} finally { $doc.Close(0) }
} finally {
	if ($word) { try { $word.Quit(0) } catch { } }
}
@{ word = $build; cases = @($cases.Keys); rows = $rows } | ConvertTo-Json -Depth 6 -Compress | Set-Content -Encoding utf8 $Out
"wrote $Out"
