param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-legend-layout'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null; $paths = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    foreach ($entry in @(@(1,1,'Month'),@(1,2,'Sales'),@(1,3,'Costs'),@(1,4,'Profit'),@(2,1,'A'),@(3,1,'B'))) { $sheet.Cells.Item($entry[0],$entry[1]).Value2 = [string]$entry[2] }
    foreach ($entry in @(@(2,2,10),@(2,3,8),@(2,4,2),@(3,2,20),@(3,3,12),@(3,4,8))) { $sheet.Cells.Item($entry[0],$entry[1]).Value2 = [double]$entry[2] }
    foreach ($name in @('right-auto','right-moved','right-overlay','bottom-auto','bottom-moved','bottom-overlay','wide-overlay','tall-overlay')) {
        $object = $sheet.ChartObjects().Add(20,20,480,300); $chart = $object.Chart
        $chart.ChartType = 51; $chart.SetSourceData($sheet.Range('A1:D3'),2)
        $chart.HasTitle = $true; $chart.ChartTitle.Text = 'Revenue'; $chart.HasLegend = $true
        $chart.Legend.Position = if ($name.StartsWith('bottom')) { -4107 } else { -4152 }
        $chart.Legend.Font.Name = 'Arial'; $chart.Legend.Font.Size = 12
        $chart.Legend.Font.Bold = $false; $chart.Legend.Font.Italic = $false
        $chart.Legend.Font.Color = 255
        $chart.Legend.Format.Fill.Visible = -1; $chart.Legend.Format.Fill.Solid(); $chart.Legend.Format.Fill.ForeColor.RGB = 65535
        if (-not $name.EndsWith('auto')) {
            $chart.Legend.IncludeInLayout = -not $name.EndsWith('overlay')
            $chart.Legend.Left = 40; $chart.Legend.Top = 100
            if ($name.StartsWith('wide')) { $chart.Legend.Width = 240; $chart.Legend.Height = 30 }
            if ($name.StartsWith('tall')) { $chart.Legend.Width = 90; $chart.Legend.Height = 150 }
        }
        $chart.Parent.Activate(); $chart.Refresh()
        $path = Join-Path $OutputFolder "legend-$name.xlsx"
        $book.SaveCopyAs($path); $paths += $path
        if (-not $chart.Export((Join-Path $OutputFolder "legend-$name.png"),'PNG')) { throw 'Native chart export failed' }
        $object.Delete()
    }
} finally {
    if ($null -ne $book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($null -ne $excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
& (Join-Path $PSScriptRoot 'record-xlsx-chart-styles.ps1') -OutputFolder $OutputFolder -ReferenceFiles $paths -CaptureTitleCharacters -CaptureLayoutGeometry
