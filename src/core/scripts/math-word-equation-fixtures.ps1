# Regenerates math/__fixtures__/word-equations.json (not run in CI; needs desktop Word on Windows).
#
#   pwsh -File scripts/math-word-equation-fixtures.ps1
#
# Word builds each equation up from its linear text through its documented automation
# (Document.OMaths.Add, OMath.BuildUp) and the Office Math it writes (Range.WordOpenXML) is kept,
# so the preservation tests in math/linear-editing.test.ts start from markup Word itself saves.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$out = Join-Path $root 'math/__fixtures__/word-equations.json'
$cases = @(
	'a/b', '(a+b)/c', 'x^2', 'x_i^2', '√x', '√(3&x)', 'sin⁡x', 'lim_(n→∞)⁡a_n', '∑_(i=1)^n▒a_i',
	'∫_0^1▒〖f(x)dx〗', '(a│b)', '|x|', '■(a&b@c&d)', '"if "x', 'a ̂', 'x┬n', 'log_2⁡x', 'e^(-x)', 'km/h',
	'x^2^3', 'a_b_c'
)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$results = @()
try {
	$doc = $word.Documents.Add()
	foreach ($case in $cases) {
		$doc.Content.Delete() | Out-Null
		$doc.Content.Text = $case
		$range = $doc.Paragraphs(1).Range
		$range.MoveEnd(1, -1) | Out-Null
		$doc.OMaths.Add($range) | Out-Null
		$doc.OMaths(1).BuildUp()
		$xml = $doc.Paragraphs(1).Range.WordOpenXML
		$omml = [regex]::Match($xml, '<m:oMathPara[\s\S]*?</m:oMathPara>|<m:oMath>[\s\S]*?</m:oMath>').Value
		$results += [pscustomobject]@{ linear = $case; omml = $omml }
	}
	$doc.Close(0)
} finally {
	$word.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
$results | ConvertTo-Json -Depth 3 | Set-Content -Encoding utf8 $out
Write-Host "wrote $($results.Count) equations to $out"
