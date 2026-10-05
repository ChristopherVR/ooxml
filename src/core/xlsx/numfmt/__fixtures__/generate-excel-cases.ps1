# Regenerates excel-format-cases.json from real Excel (Windows, Excel 16, en-US regional settings).
# Usage: pwsh -File generate-excel-cases.ps1
# Each input [value, format] is written to a wide cell (numbers as a formula `=n` so Excel does not
# re-format on entry, strings with a leading apostrophe) and the displayed Range.Text is recorded.
# Formats containing a `*` fill are skipped (formatValue drops fill characters); a cell full of
# `#` is recorded as '########' (OVERFLOW_TEXT).
$ErrorActionPreference = 'Stop'
$inputs = Get-Content -Raw -Encoding UTF8 "$PSScriptRoot\excel-format-inputs.json" | ConvertFrom-Json
$dates1904 = @(@(0, 'm/d/yyyy'), @(1, 'm/d/yyyy'), @(45000, 'm/d/yyyy'), @(0, 'dddd'), @(45000, 'dddd'),
	@(1000.75, 'yyyy-mm-dd h:mm AM/PM'), @(1461, 'm/d/yyyy'), @(1462, 'm/d/yyyy'))
$x = New-Object -ComObject Excel.Application
$x.Visible = $false
$x.DisplayAlerts = $false
if ($x.International(3) -ne '.') { throw 'Run with en-US number settings (decimal point).' }

# A workbook holds only about 250 custom formats, so a fresh workbook is used every 100 cases.
function Measure-Cases($cases, [bool]$date1904) {
	$out = @()
	$i = 0
	$wb = $null
	foreach ($c in $cases) {
		$v = $c[0]; $f = [string]$c[1]
		if ($f.Contains('*')) { continue }
		if ($v -isnot [string] -and $f -eq '@') { continue }
		if ($i % 100 -eq 0) {
			if ($wb) { $wb.Close($false) }
			$wb = $x.Workbooks.Add()
			$wb.Date1904 = $date1904
			$ws = $wb.Worksheets.Item(1)
			$ws.Columns.Item(1).ColumnWidth = 120
		}
		$i++
		$cell = $ws.Cells.Item($i, 1)
		try {
			$cell.NumberFormat = $f
			if ($v -is [string]) { $s = "'" + $v }
			elseif ($v -is [bool]) { $s = if ($v) { 'TRUE' } else { 'FALSE' } }
			else { $s = '=' + ([double]$v).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture) }
			$cell.Formula = [string]$s
			$text = [string]$cell.Text
		} catch { continue }
		if ($text -match '^#{20,}$') { $text = '########' }
		$row = @($v, $f, $text)
		if ($date1904) { $row += $true }
		$out += , $row
	}
	$wb.Close($false)
	return $out
}

$rows = @(Measure-Cases $inputs $false) + @(Measure-Cases $dates1904 $true)
$x.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) | Out-Null
$lines = $rows | ForEach-Object { "`t" + (ConvertTo-Json -InputObject $_ -Compress) }
"[`n" + ($lines -join ",`n") + "`n]" | Set-Content -Encoding UTF8 "$PSScriptRoot\excel-format-cases.json"
