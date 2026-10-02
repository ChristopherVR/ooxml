# Manual acceptance check for the xlsx writer (not run in CI; needs desktop Excel on Windows).
#
#   pwsh -File scripts/xlsx-excel-acceptance.ps1 [-OutDir <folder>] [-SkipBuild]
#
# -SkipBuild opens whatever .xlsx files are already in -OutDir (handy for checking a file by hand).
#
# 1. `bun scripts/xlsx-acceptance-build.ts` writes workbooks with `saveXlsx` (new workbooks,
#    untouched fixtures, edited fixtures).
# 2. Each file is opened in Excel through COM. A file fails when Excel throws (Excel refuses to
#    open a package it would have to repair when running unattended), when a repair log appears
#    in %TEMP% or the caption says "Repaired", when the sheet count differs from workbook.xml,
#    or when Excel's tables, charts, hyperlinks, comments or merges differ from manifest.json,
#    or when a cell listed in checks.json (values, formulas, dynamic-array spills of workbooks
#    built through an edit session) differs from what our engine computed.
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'xlsx-excel-acceptance'),
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.xlsx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Push-Location $root
	try {
		& bun scripts/xlsx-acceptance-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "xlsx-acceptance-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}

$manifestPath = Join-Path $OutDir 'manifest.json'
$manifest = if (Test-Path $manifestPath) { Get-Content -Raw $manifestPath | ConvertFrom-Json } else { $null }
$checksPath = Join-Path $OutDir 'checks.json'
$cellChecks = if (Test-Path $checksPath) { Get-Content -Raw $checksPath | ConvertFrom-Json } else { $null }

# Compares one checks.json entry (value, formula, dynamic-array spill) with what Excel computed.
function Test-CellCheck($wb, $check) {
	$ws = $wb.Worksheets.Item($check.sheet)
	$r = $ws.Range($check.cell)
	$where = "$($check.sheet)!$($check.cell)"
	$out = @()
	if ($null -ne $check.spill) {
		if (-not $r.HasSpill) { $out += "$where has no spill" }
		else {
			$spill = $r.SpillingToRange
			$size = "$($spill.Rows.Count)x$($spill.Columns.Count)"
			$want = "$($check.spill.rows)x$($check.spill.cols)"
			if ($size -ne $want) { $out += "$where spills $size expected $want" }
		}
	}
	if ($null -ne $check.value) {
		$want = $check.value
		$text = [string]$r.Text
		$got = $r.Value2
		if ($want -is [string] -and $want.StartsWith('#')) {
			if ($text -ne $want) { $out += "$where shows '$text' expected $want" }
		} elseif ($want -is [bool]) {
			if ($got -ne $want) { $out += "$where = '$got' expected $want" }
		} elseif ($want -is [string]) {
			if ([string]$got -ne $want) { $out += "$where = '$got' expected '$want'" }
		} else {
			$n = [double]$want
			if (-not ($got -is [double]) -or [math]::Abs($got - $n) -gt 1e-9 * [math]::Max(1, [math]::Abs($n))) {
				$out += "$where = '$got' expected $n"
			}
		}
	}
	if ($null -ne $check.formula -and [string]$r.Formula2 -ne $check.formula) {
		$out += "$where formula '$($r.Formula2)' expected '$($check.formula)'"
	}
	return $out
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Get-PackageSheetCount([string]$path) {
	$zip = [IO.Compression.ZipFile]::OpenRead($path)
	try {
		$reader = New-Object IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open())
		$xml = $reader.ReadToEnd()
		$reader.Close()
		return ([regex]::Matches($xml, '<sheet\s')).Count
	} finally { $zip.Dispose() }
}
function Get-RepairLogs { Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$failures = @()
$results = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.xlsx' | Sort-Object Name) {
		$before = @(Get-RepairLogs)
		$status = 'ok'
		$detail = ''
		$wb = $null
		try {
			# UpdateLinks 0 (never), ReadOnly; the default CorruptLoad is a normal load.
			$wb = $excel.Workbooks.Open($file.FullName, 0, $true)
			$caption = ''
			try { $caption = [string]$excel.ActiveWindow.Caption } catch { $caption = '' }
			$expected = Get-PackageSheetCount $file.FullName
			$count = $wb.Sheets.Count
			$after = @(Get-RepairLogs | Where-Object { $before -notcontains $_ })
			if ($after.Count -gt 0) {
				$status = 'REPAIRED'
				$detail = (Get-Content -Raw $after[0]) -replace '\s+', ' '
			} elseif ($caption -match 'Repaired') {
				$status = 'REPAIRED'
				$detail = $caption
			} elseif ($count -ne $expected) {
				$status = 'SHEETS'
				$detail = "Excel sees $count sheets, package declares $expected"
			} else {
				$detail = "$count sheet(s)"
				$expect = if ($manifest) { $manifest.($file.Name) } else { $null }
				$problems = @()
				for ($i = 1; $i -le $wb.Worksheets.Count; $i++) {
					$ws = $wb.Worksheets.Item($i)
					$e = $expect | Where-Object { $_.name -eq $ws.Name } | Select-Object -First 1
					if (-not $e) { continue }
					$threaded = 0
					try { $threaded = $ws.CommentsThreaded.Count } catch { $threaded = 0 }
					$merges = 0
					foreach ($cell in $ws.UsedRange.Cells) { if ($cell.MergeCells -and $cell.MergeArea.Cells.Item(1).Address() -eq $cell.Address()) { $merges++ } }
					$seen = [ordered]@{
						tables     = $ws.ListObjects.Count
						charts     = $ws.ChartObjects().Count
						hyperlinks = $ws.Hyperlinks.Count
						comments   = $ws.Comments.Count + $threaded
						merges     = $merges
					}
					$detail += "; $($ws.Name): " + (($seen.Keys | ForEach-Object { "$_=$($seen[$_])" }) -join ' ')
					foreach ($key in $seen.Keys) {
						$want = [int]$e.$key
						# A threaded comment shows up both as a thread and as its legacy note.
						$ok = if ($key -eq 'comments') { $seen[$key] -ge $want } else { $seen[$key] -eq $want }
						if (-not $ok) { $problems += "$($ws.Name).$key=$($seen[$key]) expected $want" }
					}
				}
				$fileChecks = if ($cellChecks) { $cellChecks.($file.Name) } else { $null }
				foreach ($check in @($fileChecks)) {
					if ($null -eq $check) { continue }
					try { $problems += @(Test-CellCheck $wb $check) } catch { $problems += "$($check.sheet)!$($check.cell): $($_.Exception.Message)" }
				}
				if ($fileChecks) { $detail += "; $(@($fileChecks).Count) cell checks" }
				if ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
			}
		} catch {
			$status = 'FAILED'
			$detail = $_.Exception.Message
		} finally {
			if ($wb) { $wb.Close($false) | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) }
		}
		$results += [pscustomobject]@{ File = $file.Name; Status = $status; Detail = $detail }
		if ($status -ne 'ok') { $failures += $file.Name }
	}
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
	[GC]::Collect()
}
$results | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
Write-Host "Workbooks are in $OutDir for inspection."
if ($failures.Count) {
	Write-Host "FAILED: $($failures -join ', ')" -ForegroundColor Red
	exit 1
}
Write-Host "All $($results.Count) workbooks opened in Excel without repair." -ForegroundColor Green
