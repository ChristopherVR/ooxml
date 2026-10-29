# Regenerates ets-init<N>-excel.json (N = 2, 3 or 4): Excel's FORECAST.ETS family evaluated on the series of
# ets-init<N>-series.json (written by `bun scripts/xlsx/ets-init2-probes.ts` in src/core).
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/core/xlsx/formula/__fixtures__/ets-init2-cases.ps1 -Set 3
# Manual step (Windows, Excel 16+): CI never runs it; the committed JSON is what tests compare with.
param([ValidateSet(2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)][int]$Set = 2)
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$series = Get-Content -Raw -Encoding UTF8 (Join-Path $here "ets-init$Set-series.json") | ConvertFrom-Json

# Another session may be using Excel: wait for it rather than killing it.
for ($i = 0; $i -lt 60 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Seconds 30 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is still running; try again later.' }

$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
	$wb = $x.Workbooks.Add()
	$ws = $wb.Sheets(1)
	$out = @()
	foreach ($s in $series) {
		$ws.Cells.Clear() | Out-Null
		$n = $s.values.Count
		for ($i = 0; $i -lt $n; $i++) {
			if ($null -ne $s.values[$i]) { $ws.Cells($i + 1, 1).Value2 = [double]$s.values[$i] }
			$ws.Cells($i + 1, 2).Value2 = [double]$s.timeline[$i]
		}
		$V = "`$A`$1:`$A`$$n"
		$L = "`$B`$1:`$B`$$n"
		$row = 1
		foreach ($f in $s.formulas) {
			$text = '=' + $f.Replace('{V}', $V).Replace('{L}', $L)
			$cell = $ws.Cells($row, 4)
			try { $cell.Formula2 = $text } catch { $cell.Value2 = 'REJECTED' }
			$row++
		}
		$x.CalculateFull()
		$results = @()
		$row = 1
		foreach ($f in $s.formulas) {
			$cell = $ws.Cells($row, 4)
			$v = $cell.Value2
			$isError = $ws.Evaluate("ISERROR(D$row)")
			if ($v -is [string] -and $v -eq 'REJECTED') { $r = @{ formula = $f; type = 'rejected' } }
			elseif ($isError -eq $true) { $r = @{ formula = $f; type = 'error'; value = $cell.Text } }
			elseif ($v -is [double]) { $r = @{ formula = $f; type = 'number'; value = $v } }
			else { $r = @{ formula = $f; type = 'string'; value = [string]$v } }
			$results += $r
			$row++
		}
		$out += [ordered]@{ name = $s.name; values = $s.values; timeline = $s.timeline; results = $results }
	}
	$json = ConvertTo-Json -InputObject $out -Depth 8
	[System.IO.File]::WriteAllText((Join-Path $here "ets-init$Set-excel.json"), $json, (New-Object System.Text.UTF8Encoding($false)))
	$wb.Close($false)
} finally {
	$x.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
