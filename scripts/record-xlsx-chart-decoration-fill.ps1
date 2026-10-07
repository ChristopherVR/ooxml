param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-decoration-fill'))
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(); $sheet = $book.Worksheets.Item(1)
    $sheet.Cells.Item(1,1).Value2 = 'Month'; $sheet.Cells.Item(1,2).Value2 = 'Sales'
    $sheet.Cells.Item(2,1).Value2 = 'A'; $sheet.Cells.Item(2,2).Value2 = [double]10
    $sheet.Cells.Item(3,1).Value2 = 'B'; $sheet.Cells.Item(3,2).Value2 = [double]20
    $chart = $sheet.ChartObjects().Add(20,20,450,300).Chart
    $chart.ChartType = 51; $chart.SetSourceData($sheet.Range('A1:B3'),2)
    $chart.HasTitle = $true; $chart.ChartTitle.Text = 'Sales chart'; $chart.HasLegend = $true
    $chart.ChartTitle.Format.Fill.Solid(); $chart.ChartTitle.Format.Fill.ForeColor.RGB = 255
    $chart.ChartTitle.Format.Fill.Transparency = 0.37
    $chart.Legend.Format.Fill.TwoColorGradient(1,1)
    $chart.Legend.Format.Fill.ForeColor.RGB = 65280
    $chart.Legend.Format.Fill.BackColor.RGB = 16777215
    $cases = @()
    foreach ($position in @(@('l',-4131),@('r',-4152),@('t',-4160),@('b',-4107),@('tr',2))) {
        $chart.Legend.Position = [int]$position[1]
        $path = Join-Path $OutputFolder ($position[0] + '.xlsx')
        $book.SaveCopyAs($path)
        $reopened = $excel.Workbooks.Open($path,0,$true)
        try {
            $native = $reopened.Worksheets.Item(1).ChartObjects(1).Chart
            $native.Parent.Activate(); $native.Refresh()
            $png = Join-Path $OutputFolder ($position[0] + '.png')
            if (-not $native.Export($png,'PNG') -or (Get-Item $png).Length -eq 0) {
                throw 'Native chart export failed'
            }
            $cases += @{
                position = $position[0]
                title = @{left=[double]$native.ChartTitle.Left;top=[double]$native.ChartTitle.Top;
                    width=[double]$native.ChartTitle.Width;height=[double]$native.ChartTitle.Height;
					fontSize=[double]$native.ChartTitle.Font.Size;fontName=[string]$native.ChartTitle.Font.Name;
                    rgb=[int]$native.ChartTitle.Format.Fill.ForeColor.RGB;
                    transparency=[double]$native.ChartTitle.Format.Fill.Transparency}
                legend = @{left=[double]$native.Legend.Left;top=[double]$native.Legend.Top;
                    width=[double]$native.Legend.Width;height=[double]$native.Legend.Height;
					fontSize=[double]$native.Legend.Font.Size;fontName=[string]$native.Legend.Font.Name;
                    foreground=[int]$native.Legend.Format.Fill.ForeColor.RGB;
                    background=[int]$native.Legend.Format.Fill.BackColor.RGB;
                    gradientStyle=[int]$native.Legend.Format.Fill.GradientStyle}
            }
        } finally {
            $reopened.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($reopened)
        }
    }
    @{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;cases=$cases} |
        ConvertTo-Json -Depth 6 | Set-Content -Encoding utf8 (Join-Path $OutputFolder 'decoration-fills.json')
    Write-Output "Recorded native decoration fills in $OutputFolder"
} finally {
    if ($book) { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
}
