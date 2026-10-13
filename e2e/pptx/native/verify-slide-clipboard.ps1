param([Parameter(Mandatory=$true)][string]$ArtifactDirectory)
$ErrorActionPreference = 'Stop'
$nativeDir = (Get-Item -LiteralPath $ArtifactDirectory).FullName
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw 'Close existing PowerPoint processes before this isolated acceptance check.'
}
$app = $null
$results = [System.Collections.Generic.List[object]]::new()

function Read-NativePresentation([string]$Path, [string]$SaveCopyPath = '') {
    $presentation = $null
    try {
        $presentation = $app.Presentations.Open((Get-Item -LiteralPath $Path).FullName, -1, 0, 0)
        $slides = @()
        for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
            $slide = $presentation.Slides.Item($i)
            $texts = @(); $colors = @(); $pictures = 0; $notes = @()
            try {
                for ($j = 1; $j -le $slide.Shapes.Count; $j++) {
                    $shape = $slide.Shapes.Item($j)
                    try {
                        if ($shape.Type -eq 13) { $pictures++ }
                        if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                            $texts += $shape.TextFrame.TextRange.Text.Trim()
                            $colors += [int]$shape.TextFrame.TextRange.Font.Color.RGB
                        }
                    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shape) }
                }
                $notesPage = $slide.NotesPage
                try {
                    for ($j = 1; $j -le $notesPage.Shapes.Count; $j++) {
                        $shape = $notesPage.Shapes.Item($j)
                        try {
                            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                                $notes += @{ text = $shape.TextFrame.TextRange.Text.Trim(); bold = [int]$shape.TextFrame.TextRange.Font.Bold }
                            }
                        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shape) }
                    }
                } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($notesPage) }
                $slides += @{ index = $i; id = $slide.SlideID; texts = $texts; colors = $colors; pictures = $pictures; notes = $notes }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($slide) }
        }
        if ($SaveCopyPath) { $presentation.SaveCopyAs([IO.Path]::GetFullPath($SaveCopyPath), 24) }
        return @{ file = [IO.Path]::GetFileName($Path); slideCount = $presentation.Slides.Count; slides = $slides }
    } finally {
        if ($presentation) { $presentation.Close(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    }
}

try {
    $app = New-Object -ComObject PowerPoint.Application
    $officeVersion = $app.Version
    foreach ($variant in @('standard', 'boundary-names')) {
        $folder = Join-Path $nativeDir $variant
        foreach ($name in @('source', 'destination', 'clipboard', 'imported', 'core-resaved', 'native-resaved')) {
            $path = Join-Path $folder "$name.pptx"
            Write-Host "Checking $variant/$name"
            $saveCopyPath = if ($name -eq 'core-resaved') { Join-Path $folder 'native-resaved.pptx' } else { '' }
            $result = Read-NativePresentation $path $saveCopyPath
            $result.variant = $variant
            $expectedCount = if ($name -eq 'clipboard') { 1 } elseif ($name -in @('source', 'destination')) { 2 } else { 3 }
            if ($result.slideCount -ne $expectedCount) { throw "$variant/$name has an unexpected slide count" }
            if ($expectedCount -eq 3) {
                $middle = $result.slides[1]
                if ($result.slides[0].texts[0] -ne 'Destination first' -or $middle.texts[0] -ne 'Source copied' -or $result.slides[2].texts[0] -ne 'Destination last') { throw "$variant/$name has an unexpected slide order" }
                if ($middle.pictures -ne 1) { throw "$variant/$name lost the source picture" }
                if ($result.slides[0].colors[0] -ne 3351057 -or $middle.colors[0] -ne 15715755) { throw "$variant/$name lost independent theme colors" }
                if (-not ($middle.notes | Where-Object { $_.text -eq 'Source rich speaker notes' -and $_.bold -eq -1 })) { throw "$variant/$name lost bold source speaker notes" }
            }
            $results.Add($result)
        }
    }
    @{ checkedAtUtc = [DateTime]::UtcNow.ToString('o'); officeVersion = $officeVersion; cases = $results } |
        ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $nativeDir 'native-acceptance.json')
    Write-Output "Passed $($results.Count) native PowerPoint open/save/reopen checks."
} finally {
    if ($app) {
        if ($app.Presentations.Count -eq 0) { $app.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    }
}
