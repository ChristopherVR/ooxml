# Export the generated Word layout issue repros via desktop Word on Windows.
# Pass -InputDir to check saved round trips with the same tool.
param(
	[string]$InputDir = (Join-Path $PSScriptRoot '../docx/layout/fixtures/word-layout-reports'),
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-layout-reports-word')
)
$ErrorActionPreference = 'Stop'
$files = @(Get-ChildItem -LiteralPath $InputDir -Filter '*.docx' -File | Sort-Object Name)
if ($files.Count -eq 0) { throw "No DOCX files in $InputDir" }
$null = New-Item -ItemType Directory -Force $OutDir
$output = (Resolve-Path -LiteralPath $OutDir).Path

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.AutomationSecurity = 3
try {
	Write-Host "Word $($word.Version), build $($word.Build)"
	foreach ($file in $files) {
		$doc = $null
		try {
			$doc = $word.Documents.Open($file.FullName, $false, $true, $false)
			if ([string]$doc.ActiveWindow.Caption -match 'Repaired') {
				throw "Word repaired $($file.Name)"
			}
			$doc.Repaginate()
			$doc.ExportAsFixedFormat((Join-Path $output "$($file.BaseName).pdf"), 17)
			[pscustomobject]@{
				File = $file.Name
				Pages = $doc.ComputeStatistics(2)
			}
		} finally {
			if ($doc) {
				$doc.Close(0)
				[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
			}
		}
	}
} finally {
	$word.Quit(0)
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
