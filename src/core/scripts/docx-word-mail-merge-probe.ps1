# Opens merged documents (`mergeFields` output) in desktop Word on Windows through its documented
# automation and prints what Word shows: whether it repaired the file, the text, tables and
# lists, and, after Update Field (F9), each field's result text and the formatting of its second
# character. The result formatting rules in docx/mail-merge/result-runs.ts and the expectations
# in docx/mail-merge/headless.test.ts were read this way. Manual; run from any folder:
#   powershell -ExecutionPolicy Bypass -File docx-word-mail-merge-probe.ps1 -InputDir <folder>
param([Parameter(Mandatory = $true)][string]$InputDir)
$files = @(Get-ChildItem -LiteralPath $InputDir -Filter '*.docx' -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No DOCX files in $InputDir" }
$visible = { param($text) $text -replace "`r", '\r' -replace "$([char]11)", '\v' -replace "`t", '\t' -replace "$([char]7)", '\a' }
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$started = @(Get-Process WINWORD | Where-Object { $before -notcontains $_.Id } | ForEach-Object Id)
try {
	foreach ($file in $files) {
		$doc = $word.Documents.Open($file.FullName, $false, $true)
		"== $($file.Name): repaired=$([string]$doc.ActiveWindow.Caption -match 'Repaired') tables=$($doc.Tables.Count) lists=$($doc.Lists.Count) fields=$($doc.Fields.Count)"
		& $visible $doc.Content.Text
		foreach ($table in $doc.Tables) { "table rows=$($table.Rows.Count) borders=$($table.Borders.Enable)" }
		foreach ($field in $doc.Fields) {
			[void]$field.Update()
			$result = $field.Result
			$char = $result.Characters.Item([Math]::Min(2, [Math]::Max(1, $result.Characters.Count)))
			'{0} => "{1}" bold={2} italic={3} underline={4}' -f $field.Code.Text, (& $visible $result.Text), $char.Bold, $char.Italic, $char.Underline
		}
		$doc.Close([ref]0)
	}
} finally {
	$word.Quit([ref]0)
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
	Start-Sleep 2
	foreach ($id in $started) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
