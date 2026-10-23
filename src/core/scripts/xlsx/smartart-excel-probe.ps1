# Probe (manual, Windows, desktop Excel): how Excel treats a SmartArt whose data part or cached
# drawing changed behind its back. Evidence for docs/agnostic-core-plan.md ("Excel and an edited
# data part"); the Excel counterpart of scripts/docx/smartart-word-probe.ps1.
#
#   pwsh -File scripts/xlsx/smartart-excel-probe.ps1 [-OutDir <folder>]
#
# Each variant built by smartart-excel-variants.mjs from xlsx/__fixtures__/excel-smartart.xlsx is
# opened read-only; the script prints the node texts Excel reports, which probe words appear in the
# picture Excel draws (the shape copied to the clipboard as an EMF), and the shapes and texts of the
# cached drawing Excel writes when it saves the workbook again.
param([string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'xlsx-smartart-probe'))
$ErrorActionPreference = 'Stop'
$core = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))
$fixture = Join-Path $core 'xlsx/__fixtures__/excel-smartart.xlsx'
if (Test-Path $OutDir) { Remove-Item -Recurse -Force $OutDir }
New-Item -ItemType Directory -Force $OutDir | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class EmfClipboard {
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

Push-Location $core
try { & bun scripts/xlsx/smartart-excel-variants.mjs $fixture $OutDir; if ($LASTEXITCODE) { throw 'variants failed' } }
finally { Pop-Location }
Copy-Item $fixture (Join-Path $OutDir 'original.xlsx')

$probeWords = @('Plan', 'Build', 'EDITED', 'Ship', 'Cacheonly', 'Zeta')
$resaved = Join-Path $OutDir 'resaved'
New-Item -ItemType Directory -Force $resaved | Out-Null
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.xlsx' | Sort-Object Name) {
		$wb = $null
		try {
			$wb = $excel.Workbooks.Open($file.FullName, 0, $true)
			$shape = $wb.Worksheets.Item(1).Shapes.Item(1)
			$nodes = $shape.SmartArt.AllNodes
			$texts = @(); for ($i = 1; $i -le $nodes.Count; $i++) { $texts += $nodes.Item($i).TextFrame2.TextRange.Text }
			# xlScreen, xlPicture: an enhanced metafile on the clipboard, text as EMR_EXTTEXTOUTW.
			$shape.CopyPicture(1, -4147)
			$emf = [Text.Encoding]::Unicode.GetString([EmfClipboard]::Read())
			$drawn = $probeWords | Where-Object { $emf.Contains($_) }
			$copy = Join-Path $resaved $file.Name
			$wb.SaveAs($copy, 51)
			$wb.Close($false); $wb = $null
			$zip = [IO.Compression.ZipFile]::OpenRead($copy)
			try {
				$entry = $zip.Entries | Where-Object { $_.FullName -like 'xl/diagrams/drawing*.xml' } | Select-Object -First 1
				$xml = if ($entry) { (New-Object IO.StreamReader($entry.Open())).ReadToEnd() } else { '' }
			} finally { $zip.Dispose() }
			$shapes = ([regex]::Matches($xml, '<dsp:sp ')).Count
			$saved = ([regex]::Matches($xml, '<a:t>([^<]*)') | ForEach-Object { $_.Groups[1].Value }) -join '|'
			Write-Host "$($file.BaseName): nodes [$($texts -join ', ')]; drawn [$($drawn -join ', ')]; re-saved drawing $shapes shapes [$saved]"
		} catch {
			Write-Host "$($file.BaseName): FAILED $($_.Exception.Message)" -ForegroundColor Red
		} finally { if ($wb) { $wb.Close($false) | Out-Null } }
	}
} finally {
	$excel.Quit()
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
	[GC]::Collect()
}
