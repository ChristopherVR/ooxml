param([string]$OutputFolder = (Join-Path $env:TEMP 'ooxml-native-chart-spacing'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
$excel=$null; $book=$null
function Read-Capture($application,$path,$type){
    $zip=[System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $parts=@{}
        foreach($entry in $zip.Entries){if($entry.FullName -notmatch '^xl/charts/'){continue}; $reader=[System.IO.StreamReader]::new($entry.Open()); try{$parts[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}}
    }finally{$zip.Dispose()}
    $probe=$application.Workbooks.Open($path,0,$true)
    try {
        $native=$probe.Worksheets.Item(1).ChartObjects(1).Chart
        $points=@()
        foreach($s in 1..3){foreach($p in 1..2){$point=$native.SeriesCollection($s).Points($p); $points+=@{series=$s-1;point=$p-1;left=[double]$point.Left;top=[double]$point.Top;width=[double]$point.Width;height=[double]$point.Height}}}
        return @{type=$type;gapWidth=[int]$native.ChartGroups(1).GapWidth;overlap=[int]$native.ChartGroups(1).Overlap;points=$points;parts=$parts}
    }finally{$probe.Close($false)}
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false
    $book=$excel.Workbooks.Add(); $sheet=$book.Worksheets.Item(1)
    foreach($col in 1..4){$sheet.Cells.Item(1,$col).Value2=[string]"C$col"}
    foreach($row in 2..5){
        $sheet.Cells.Item($row,1).Value2=[string]"M$row"
        foreach($col in 2..4){$sheet.Cells.Item($row,$col).Value2=[double]($row*$col)}
    }
    $scheme=@{}; $names=@('dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink')
    for($i=0;$i -lt $names.Count;$i++){$rgb=[int]$book.Theme.ThemeColorScheme.Colors($i+1).RGB; $scheme[$names[$i]]='#{0:X2}{1:X2}{2:X2}' -f ($rgb -band 255),(($rgb -shr 8) -band 255),(($rgb -shr 16) -band 255)}
    $cases=@(); $options=@(@(100,-24),@(5,23),@(219,-27),@(0,100),@(500,-100))
    foreach($type in @(51,57,52,53)){
        $selectedOptions=$options
        if($type -in @(52,53)){$selectedOptions=@(,@(100,100))}
        foreach($option in $selectedOptions){
            $object=$sheet.ChartObjects().Add(260,20,480,300); $chart=$object.Chart
            $chart.ChartType=$type; $chart.SetSourceData($sheet.Range('A1:D5'),2)
            $chart.ChartStyle=201; $chart.ChartColor=10
            $chart.ChartGroups(1).GapWidth=[int]$option[0]; $chart.ChartGroups(1).Overlap=[int]$option[1]
            $path=Join-Path $OutputFolder "spacing-$type-$($option[0])-$($option[1]).xlsx"
            $book.SaveCopyAs($path)
            $cases+=Read-Capture $excel $path $type
            $object.Delete()
        }
    }
    # Remove only the overlap declaration from an owned saved workbook to measure Excel's default.
    $defaultPath=Join-Path $OutputFolder 'stacked-default.xlsx'
    Copy-Item -LiteralPath (Join-Path $OutputFolder 'spacing-52-100-100.xlsx') -Destination $defaultPath -Force
    $zip=[System.IO.Compression.ZipFile]::Open($defaultPath,[System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry=$zip.GetEntry('xl/charts/chart1.xml'); $reader=[System.IO.StreamReader]::new($entry.Open())
        try{$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
        $pattern='<c:overlap[^>]*/>'
        if([regex]::Matches($xml,$pattern).Count -ne 1){throw 'Expected one native overlap declaration'}
        $xml=[regex]::Replace($xml,$pattern,''); $entry.Delete()
        $writer=[System.IO.StreamWriter]::new($zip.CreateEntry('xl/charts/chart1.xml').Open())
        try{$writer.Write($xml)}finally{$writer.Dispose()}
    }finally{$zip.Dispose()}
    $cases+=Read-Capture $excel $defaultPath 52
    @{excelVersion=[string]$excel.Version;excelBuild=[string]$excel.Build;scheme=$scheme;cases=$cases}|ConvertTo-Json -Depth 9|Set-Content -Encoding utf8 (Join-Path $OutputFolder 'spacing.json')
    Write-Output "Recorded $($cases.Count) native spacing cases in $OutputFolder"
}finally{if($book){$book.Close($false)};if($excel){$excel.Quit()}}
