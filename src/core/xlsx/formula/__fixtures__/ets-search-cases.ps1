# Regenerates ets-search-excel.json: Excel's FORECAST.ETS answers on the series of
# ets-search-series.json (written by `bun scripts/xlsx/ets-search-series.ts` in src/core), used to
# identify how Excel searches for alpha, beta and gamma. All series share one sheet and one
# calculation, so a session is short.
#   powershell -NoProfile -ExecutionPolicy Bypass -File src/core/xlsx/formula/__fixtures__/ets-search-cases.ps1
# Manual step (Windows, Excel 16+): CI never runs it; the committed JSON is what tests compare with.
param([string]$In = 'ets-search-series.json', [string]$Out = (Join-Path $env:TEMP 'ets-search-excel.json'))
# -Out is a full path (a deep worktree path can exceed Windows' limit for .NET file calls); copy the
# result over ets-search-excel.json next to this script.
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$series = Get-Content -Raw -Encoding UTF8 (Join-Path $here $In) | ConvertFrom-Json

# Another session may be using Excel: wait for it rather than killing it.
for ($i = 0; $i -lt 60 -and (Get-Process EXCEL -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Seconds 30 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Excel is still running; try again later.' }

$x = New-Object -ComObject Excel.Application
$x.DisplayAlerts = $false
try {
	$wb = $x.Workbooks.Add()
	$ws = $wb.Sheets(1)
	$col = 1
	$cells = @()
	foreach ($s in $series) {
		$n = $s.values.Count
		$va = New-Object 'object[,]' $n, 1
		$la = New-Object 'object[,]' $n, 1
		for ($i = 0; $i -lt $n; $i++) {
			if ($null -ne $s.values[$i]) { $va[$i, 0] = [double]$s.values[$i] }
			$la[$i, 0] = [double]$s.timeline[$i]
		}
		$ws.Range($ws.Cells(1, $col), $ws.Cells($n, $col)).Value2 = $va
		$ws.Range($ws.Cells(1, $col + 1), $ws.Cells($n, $col + 1)).Value2 = $la
		$vl = $ws.Range($ws.Cells(1, $col), $ws.Cells($n, $col)).Address($true, $true)
		$ll = $ws.Range($ws.Cells(1, $col + 1), $ws.Cells($n, $col + 1)).Address($true, $true)
		$row = 1
		$fc = $col + 2
		$list = @()
		foreach ($f in $s.formulas) {
			$text = '=' + $f.Replace('{V}', $vl).Replace('{L}', $ll)
			$cell = $ws.Cells($row, $fc)
			try { $cell.Formula2 = $text } catch { $cell.Value2 = 'REJECTED' }
			$list += $cell
			$row++
		}
		$cells += , $list
		$col += 4
	}
	$x.CalculateFull()
	$rows = @()
	for ($k = 0; $k -lt $series.Count; $k++) {
		$s = $series[$k]
		$results = @()
		for ($j = 0; $j -lt $s.formulas.Count; $j++) {
			$cell = $cells[$k][$j]
			$f = $s.formulas[$j]
			$v = $cell.Value2
			if ($v -is [string] -and $v -eq 'REJECTED') { $r = [ordered]@{ formula = $f; type = 'rejected' } }
			elseif ($v -is [int] -and $cell.Text.StartsWith('#')) { $r = [ordered]@{ formula = $f; type = 'error'; value = $cell.Text } }
			elseif ($v -is [double]) { $r = [ordered]@{ formula = $f; type = 'number'; value = $v } }
			else { $r = [ordered]@{ formula = $f; type = 'string'; value = [string]$v } }
			$results += $r
		}
		$rows += [ordered]@{ name = $s.name; values = $s.values; timeline = $s.timeline; results = $results }
	}
	$lines = $rows | ForEach-Object { "`t" + (ConvertTo-Json -InputObject $_ -Depth 6 -Compress) }
	$json = "[`n" + ($lines -join ",`n") + "`n]`n"
	[System.IO.File]::WriteAllText($Out, $json, (New-Object System.Text.UTF8Encoding($false)))
	$wb.Close($false)
} finally {
	$x.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
}
