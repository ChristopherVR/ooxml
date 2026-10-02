# Regenerates excel-results.json by evaluating excel-cases.txt in real Excel (Windows, Excel 16+).
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/xlsx/formula/__fixtures__/excel-cases.ps1
# Manual step: CI never runs it; the committed JSON is what the tests compare against.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$lines = Get-Content -Encoding UTF8 (Join-Path $here 'excel-cases.txt')
$x = New-Object -ComObject Excel.Application
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
		if ($v -is [string] -and $v -eq 'REJECTED') { $entry = @{ formula = $f; type = 'rejected' } }
		elseif ($isError -eq $true) { $entry = @{ formula = $f; type = 'error'; value = $cell.Text } }
		elseif ($v -is [bool]) { $entry = @{ formula = $f; type = 'boolean'; value = $v } }
		elseif ($v -is [double]) { $entry = @{ formula = $f; type = 'number'; value = $v } }
		elseif ($null -eq $v) { $entry = @{ formula = $f; type = 'empty' } }
		else { $entry = @{ formula = $f; type = 'string'; value = [string]$v } }
		$out += $entry
		$row++
	}
	$json = "[`n" + (($out | ForEach-Object { '  ' + (ConvertTo-Json -InputObject $_ -Compress) }) -join ",`n") + "`n]`n"
	[System.IO.File]::WriteAllText((Join-Path $here 'excel-results.json'), $json, (New-Object System.Text.UTF8Encoding($false)))
	$wb.Close($false)
} finally {
	$x.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
