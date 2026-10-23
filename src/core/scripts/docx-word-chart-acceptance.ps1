# Manual acceptance check for Word chart editing (not run in CI; needs desktop Word on Windows).
#
#   pwsh -File scripts/docx-word-chart-acceptance.ps1 [-OutDir <folder>] [-SkipBuild] [-SkipWorkbook]
#
# 1. Word itself writes `word-source.docx`: a document with a chart from Word's Insert Chart.
# 2. `bun scripts/docx-word-chart-build.ts` saves documents with `saveDocx` after chart edits (a
#    changed value; renamed, added and relabelled series and categories; a type change; new
#    charts in a blank document and beside Word's own chart), on Word's chart and on the
#    Excel-authored fixture, and writes manifest.json with what each chart must show.
# 3. Each document is opened in Word through COM with OpenAndRepair off and alerts off, so a
#    package Word would have to repair throws. It is opened read-write (a read-only document
#    cannot open its chart data) and closed without saving. A file fails when it throws, when
#    Word sees a different number of charts, or when a chart's series names or values differ
#    from the manifest. Unless -SkipWorkbook, each chart's embedded workbook is opened with
#    ChartData.ActivateChartDataWindow and the manifest's cells are compared, so Edit Data in
#    Word shows the edited numbers (Word refuses a workbook whose table no longer matches its
#    header cells).
param(
	[string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'docx-word-chart-acceptance'),
	[switch]$SkipBuild,
	[switch]$SkipWorkbook
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

# Kills the Word and Excel processes this script started, and only those.
$preexisting = @(Get-Process WINWORD, EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
function Stop-StartedOffice {
	# CIM Terminate also ends an automation instance Stop-Process reports as already gone.
	Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='EXCEL.EXE'" |
		Where-Object { $preexisting -notcontains $_.ProcessId } |
		ForEach-Object { try { [void](Invoke-CimMethod -InputObject $_ -MethodName Terminate) } catch { } }
}

function New-Word {
	$app = New-Object -ComObject Word.Application
	$app.Visible = $false
	$app.DisplayAlerts = 0 # wdAlertsNone
	return $app
}

# Documents.Open with OpenAndRepair off; Visible is left at its default because a document
# opened invisible reports no charts.
function Open-Document($app, [string]$path) {
	# Named arguments through InvokeMember: PowerShell 7 cannot pass COM's "missing" positionally.
	$docs = $app.Documents
	$names = [string[]]@('FileName', 'ConfirmConversions', 'ReadOnly', 'AddToRecentFiles', 'OpenAndRepair', 'NoEncodingDialog')
	$values = [object[]]@($path, $false, $false, $false, $false, $true)
	# The comma keeps PowerShell from enumerating the returned document.
	return , [System.__ComObject].InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $docs, $values, $null, $null, $names)
}

if (-not $SkipBuild) {
	if (Test-Path $OutDir) { Get-ChildItem $OutDir -Filter '*.docx' | Remove-Item -Force }
	New-Item -ItemType Directory -Force $OutDir | Out-Null
	$word = New-Word
	try {
		$doc = $word.Documents.Add()
		$doc.Content.Text = 'A chart inserted by Word.'
		$range = $doc.Content
		$range.Collapse(0) # wdCollapseEnd
		$shape = $doc.InlineShapes.AddChart2(-1, 51, $range) # xlColumnClustered
		try { $shape.Chart.ChartData.Workbook.Close() } catch { }
		$doc.SaveAs2((Join-Path $OutDir 'word-source.docx'), 16) # wdFormatDocumentDefault
		$doc.Close($false)
	} finally {
		$word.Quit($false)
		[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
		Stop-StartedOffice
	}
	Push-Location $root
	try {
		& bun scripts/docx-word-chart-build.ts $OutDir
		if ($LASTEXITCODE -ne 0) { throw "docx-word-chart-build.ts failed ($LASTEXITCODE)" }
	} finally { Pop-Location }
}

$manifest = Get-Content -Raw (Join-Path $OutDir 'manifest.json') | ConvertFrom-Json

function Test-Number($got, $want) {
	if ($null -eq $want) { return ($null -eq $got) -or ([string]$got -eq '') }
	$n = 0.0
	if (-not [double]::TryParse([string]$got, [ref]$n)) { return $false }
	return [math]::Abs($n - [double]$want) -le 1e-9 * [math]::Max(1, [math]::Abs([double]$want))
}

# What Word reads from one chart, compared with its manifest entry.
function Test-Chart($chart, $want, [string]$where) {
	$problems = @()
	$count = $chart.SeriesCollection().Count
	if ($count -ne @($want.series).Count) { return @("$where has $count series, expected $(@($want.series).Count)") }
	for ($s = 0; $s -lt $count; $s++) {
		$series = $chart.SeriesCollection($s + 1)
		$expected = $want.series[$s]
		if ([string]$series.Name -ne $expected.name) { $problems += "$where series $($s + 1) is '$($series.Name)', expected '$($expected.name)'" }
		$values = @($series.Values)
		$wanted = @($expected.values)
		if ($values.Count -ne $wanted.Count) { $problems += "$where series $($s + 1) has $($values.Count) values, expected $($wanted.Count)"; continue }
		for ($i = 0; $i -lt $values.Count; $i++) {
			if (-not (Test-Number $values[$i] $wanted[$i])) { $problems += "$where series $($s + 1) value $($i + 1) is '$($values[$i])', expected '$($wanted[$i])'" }
		}
	}
	if (-not $SkipWorkbook) {
		$data = $chart.ChartData
		try {
			$data.ActivateChartDataWindow()
			$sheet = $data.Workbook.Worksheets.Item(1)
			foreach ($cell in $want.cells.PSObject.Properties) {
				$text = [string]$sheet.Range($cell.Name).Value2
				$same = ($text -eq $cell.Value) -or (Test-Number $text $cell.Value)
				if (-not $same) { $problems += "$where workbook $($cell.Name) is '$text', expected '$($cell.Value)'" }
			}
		} catch {
			$problems += "$where workbook could not be opened: $($_.Exception.Message)"
		} finally {
			try { $data.Workbook.Close() } catch { }
		}
	}
	return $problems
}

$word = New-Word
$results = @()
$failures = @()
try {
	foreach ($file in Get-ChildItem $OutDir -Filter '*.docx' | Sort-Object Name) {
		$status = 'ok'
		$detail = ''
		$doc = $null
		try {
			$doc = Open-Document $word $file.FullName
			$charts = [Collections.Generic.List[object]]::new()
			for ($i = 1; $i -le $doc.InlineShapes.Count; $i++) {
				$shape = $doc.InlineShapes.Item($i)
				if ($shape.HasChart -eq -1 -or $shape.HasChart -eq $true) { $charts.Add($shape.Chart) }
			}
			$want = @($manifest.($file.Name))
			if ($file.Name -eq 'word-source.docx') { $detail = "$($charts.Count) chart(s) written by Word" }
			elseif ($charts.Count -ne $want.Count) { $status = 'CHARTS'; $detail = "Word sees $($charts.Count) charts, expected $($want.Count)" }
			else {
				$problems = @()
				for ($c = 0; $c -lt $charts.Count; $c++) { $problems += @(Test-Chart $charts[$c] $want[$c] "chart $($c + 1)") }
				if ($problems.Count) { $status = 'CONTENT'; $detail = $problems -join '; ' }
				else { $detail = "$($charts.Count) chart(s) match" + $(if ($SkipWorkbook) { '' } else { ', workbooks included' }) }
			}
		} catch {
			$status = 'FAILED'
			$detail = $_.Exception.Message
		} finally {
			if ($doc) { $doc.Close($false) | Out-Null; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
		}
		$results += [pscustomobject]@{ File = $file.Name; Status = $status; Detail = $detail }
		if ($status -ne 'ok') { $failures += $file.Name }
	}
} finally {
	try { $word.Quit($false) } catch { }
	[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
	[GC]::Collect()
	Stop-StartedOffice
}
$results | Format-Table -AutoSize -Wrap | Out-String -Width 220 | Write-Host
Write-Host "Documents are in $OutDir for inspection."
if ($failures.Count) {
	Write-Host "FAILED: $($failures -join ', ')" -ForegroundColor Red
	exit 1
}
Write-Host "All $($results.Count) documents opened in Word without repair." -ForegroundColor Green
