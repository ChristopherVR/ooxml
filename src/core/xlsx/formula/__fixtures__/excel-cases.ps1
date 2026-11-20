# Regenerates <corpus>-results.json by evaluating <corpus>-cases.txt in real Excel (Windows,
# Excel 16+), through documented COM automation on a hidden instance this script starts.
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/core/xlsx/formula/__fixtures__/excel-cases.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/core/xlsx/formula/__fixtures__/excel-cases.ps1 -Corpus excel-p4
# Manual step: CI never runs it; the committed JSON is what the tests compare against. Outside the
# original corpus, a number also records under `text` what the cell shows in a wide General column.
param([string]$Corpus = 'excel')
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$lines = Get-Content -Encoding UTF8 (Join-Path $here "$Corpus-cases.txt")
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
try {
	$wb = $x.Workbooks.Add()
	$ws = $wb.Sheets(1)
	$ws.Columns('Z').ColumnWidth = 120
	$cases = @()
	foreach ($l in $lines) {
		if ($l.StartsWith('#') -or $l.Trim() -eq '') { continue }
		if ($l.StartsWith('!')) {
			$eq = $l.IndexOf('=')
			$ws.Range($l.Substring(1, $eq - 1)).Formula = $l.Substring($eq + 1)
			continue
		}
		$cases += [string]$l
	}
	$row = 1
	foreach ($f in $cases) {
		$cell = $ws.Range("Z$row")
		try { $cell.Formula2 = $f } catch { $cell.Value2 = 'REJECTED' }
		$row++
	}
	$x.CalculateFull()
	$out = @()
	$row = 1
	foreach ($f in $cases) {
		$cell = $ws.Range("Z$row")
		$v = $cell.Value2
		$isError = $ws.Evaluate("ISERROR(Z$row)")
		if ($v -is [string] -and $v -eq 'REJECTED') { $entry = [ordered]@{ formula = $f; type = 'rejected' } }
		elseif ($isError -eq $true) { $entry = [ordered]@{ formula = $f; type = 'error'; value = $cell.Text } }
		elseif ($v -is [bool]) { $entry = [ordered]@{ formula = $f; type = 'boolean'; value = $v } }
		elseif ($v -is [double]) { $entry = [ordered]@{ formula = $f; type = 'number'; value = $v } }
		elseif ($null -eq $v) { $entry = [ordered]@{ formula = $f; type = 'empty' } }
		else { $entry = [ordered]@{ formula = $f; type = 'string'; value = [string]$v } }
		if ($Corpus -ne 'excel' -and $entry.type -eq 'number') { $entry.text = [string]$cell.Text }
		$out += $entry
		$row++
	}
	$json = "[`n" + (($out | ForEach-Object { '  ' + (ConvertTo-Json -InputObject $_ -Compress) }) -join ",`n") + "`n]`n"
	[System.IO.File]::WriteAllText((Join-Path $here "$Corpus-results.json"), $json, (New-Object System.Text.UTF8Encoding($false)))
	$wb.Saved = $true
	$wb.Close($false)
} finally {
	$x.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
