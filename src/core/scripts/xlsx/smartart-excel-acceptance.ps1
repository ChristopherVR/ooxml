# Manual acceptance check for Excel SmartArt editing (not run in CI; needs desktop Excel on Windows).
#
#   pwsh -File scripts/xlsx/smartart-excel-acceptance.ps1 [-OutDir <folder>] [-SkipBuild]
#
# 1. `bun scripts/xlsx/smartart-acceptance-build.ts` writes workbooks with `saveXlsx`: the Excel
#    SmartArt fixture untouched, after node edits, after promote/demote, after Change Layout, and
#    a new workbook with four presets inserted. `manifest.json` lists the expected node texts.
# 2. Each file is opened read-only in Excel through COM. A file fails when Excel throws (it refuses
#    to open a package it would have to repair when unattended), a repair log appears in %TEMP%,
#    the caption says "Repaired", a SmartArt's `AllNodes` texts or layout id differ from the
#    manifest, or a node's text is missing from the picture Excel draws (the shape copied to the
#    clipboard as an EMF).
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'xlsx-smartart-acceptance'),
	[switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
$core = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.xlsx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	Push-Location $core
	try {
		& bun scripts/xlsx/smartart-acceptance-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "smartart-acceptance-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}
$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class SmartArtEmfClipboard {
	[DllImport("user32.dll")] static extern bool OpenClipboard(IntPtr owner);
	[DllImport("user32.dll")] static extern bool CloseClipboard();
	[DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint format);
	[DllImport("gdi32.dll")] static extern uint GetEnhMetaFileBits(IntPtr emf, uint size, byte[] data);
	public static byte[] Read() {
		if (!OpenClipboard(IntPtr.Zero)) return new byte[0];
		try {
			IntPtr emf = GetClipboardData(14);
			if (emf == IntPtr.Zero) return new byte[0];
			uint size = GetEnhMetaFileBits(emf, 0, null);
			byte[] data = new byte[size];
			GetEnhMetaFileBits(emf, size, data);
			return data;
		} finally { CloseClipboard(); }
	}
}
'@
function Get-RepairLogs { Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'error*.xml' -ErrorAction SilentlyContinue | ForEach-Object FullName }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$results = @()
$failures = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.xlsx' | Sort-Object Name) {
		$before = @(Get-RepairLogs)
		$status = 'ok'
		$detail = ''
		$wb = $null
		try {
			$wb = $excel.Workbooks.Open($file.FullName, 0, $true)
			$caption = ''
			try { $caption = [string]$excel.ActiveWindow.Caption } catch { $caption = '' }
			$after = @(Get-RepairLogs | Where-Object { $before -notcontains $_ })
			if ($after.Count -gt 0) { throw "repaired: $((Get-Content -Raw $after[0]) -replace '\s+', ' ')" }
			if ($caption -match 'Repaired') { throw "repaired: $caption" }
			$arts = @()
			foreach ($ws in $wb.Worksheets) { foreach ($shape in $ws.Shapes) { if ($shape.HasSmartArt) { $arts += $shape } } }
			$want = @($manifest.($file.Name))
			if ($arts.Count -ne $want.Count) { throw "Excel sees $($arts.Count) SmartArt, expected $($want.Count)" }
			$problems = @()
			$summary = @()
			for ($i = 0; $i -lt $arts.Count; $i++) {
				$art = $arts[$i].SmartArt
				$texts = @(); for ($n = 1; $n -le $art.AllNodes.Count; $n++) { $texts += [string]$art.AllNodes.Item($n).TextFrame2.TextRange.Text }
				$expect = @($want[$i].texts)
				if (($texts -join '|') -ne ($expect -join '|')) { $problems += "SmartArt $($i + 1) nodes [$($texts -join ', ')] expected [$($expect -join ', ')]" }
				if ($want[$i].layout -and $art.Layout.Id -ne $want[$i].layout) { $problems += "SmartArt $($i + 1) layout $($art.Layout.Id) expected $($want[$i].layout)" }
				$arts[$i].CopyPicture(1, -4147)
				$emf = [Text.Encoding]::Unicode.GetString([SmartArtEmfClipboard]::Read())
				# EMF text records hold the drawn glyph runs; an ampersand-free word of each text is enough.
				$missing = $expect | Where-Object { $_ } | Where-Object { -not $emf.Contains(($_ -split ' ')[0]) }
				if ($missing) { $problems += "SmartArt $($i + 1) does not draw [$($missing -join ', ')]" }
				$summary += "[$($texts -join ', ')]"
			}
			$detail = $summary -join ' '
			if ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
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
Write-Host "All $($results.Count) workbooks opened in Excel without repair with the expected SmartArt." -ForegroundColor Green
