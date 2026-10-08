# Native overlap and tie distribution reference, using only an owned invisible instance.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-distribution-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Positions($doc) {
    @($doc.Pages | ForEach-Object {
        [ordered]@{ pageId=[string]$_.ID; name=$_.Name; shapes=@($_.Shapes | ForEach-Object {
            [ordered]@{ id=[string]$_.ID; pinX=$_.CellsU('PinX').ResultIU; pinY=$_.CellsU('PinY').ResultIU }
        }) }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-Positions $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened page count differs.' }
        for ($index = 0; $index -lt $actual.Count; $index++) {
            $expected = $reference.native[$index]
            $observed = $actual[$index]
            if ($observed.pageId -cne $expected.pageId -or $observed.shapes.Count -ne $expected.shapes.Count) {
                throw "Native reopen mismatch: $($observed.name) shape/page identity"
            }
            for ($shape = 0; $shape -lt $observed.shapes.Count; $shape++) {
                if ($observed.shapes[$shape].id -cne $expected.shapes[$shape].id) {
                    throw "Native reopen mismatch: $($observed.name) shape $shape identity"
                }
                foreach ($key in @('pinX','pinY')) {
                    if ([Math]::Abs($observed.shapes[$shape][$key] - $expected.shapes[$shape].$key) -gt 1e-8) {
                        throw "Native reopen mismatch: $($observed.name) shape $shape $key"
                    }
                }
            }
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        $scenarios = @(
            @{ name='nested'; starts=@(0,2,3); ends=@(1,100,4); order=@('1','2','3') },
            @{ name='overlap'; starts=@(0,3,4); ends=@(5,4,6); order=@('1','2','3') },
            @{ name='enclosing'; starts=@(0,2,4); ends=@(100,3,5); order=@('1','2','3') },
            @{ name='leading-tie'; starts=@(0,0,4); ends=@(1,2,5); order=@('1','2','3') },
            @{ name='leading-tie-reverse'; starts=@(0,0,4); ends=@(1,2,5); order=@('2','1','3') },
            @{ name='trailing-tie'; starts=@(0,3,4); ends=@(1,5,5); order=@('1','2','3') },
            @{ name='trailing-tie-reverse'; starts=@(0,3,4); ends=@(1,5,5); order=@('3','2','1') },
            @{ name='all-leading-tie'; starts=@(0,0,0); ends=@(1,2,3); order=@('1','2','3') },
            @{ name='all-leading-tie-reverse'; starts=@(0,0,0); ends=@(1,2,3); order=@('3','2','1') },
            @{ name='same-box'; starts=@(0,0,0); ends=@(2,2,2); order=@('1','2','3') },
            @{ name='four-overlap'; starts=@(0,2,3,4); ends=@(1,100,4,6); order=@('1','2','3','4') },
            @{ name='four-overlap-reverse'; starts=@(0,2,3,4); ends=@(1,100,4,6); order=@('4','3','2','1') },
            @{ name='center-tie'; starts=@(0,1,6); ends=@(4,3,7); order=@('1','2','3') },
            @{ name='center-tie-reverse'; starts=@(0,1,6); ends=@(4,3,7); order=@('2','1','3') },
            @{ name='center-tie-short-first'; starts=@(1,0,6); ends=@(3,4,7); order=@('1','2','3') },
            @{ name='all-center-tie'; starts=@(0,1,0.5); ends=@(4,3,3.5); order=@('1','2','3') },
            @{ name='all-center-tie-reverse'; starts=@(0,1,0.5); ends=@(4,3,3.5); order=@('3','2','1') }
        )
        $cases = @()
        foreach ($scenario in $scenarios) {
            foreach ($axis in @('horizontal','vertical')) {
                $page = if ($document.Pages.Item(1).Shapes.Count -eq 0) { $document.Pages.Item(1) } else { $document.Pages.Add() }
                $page.Name = "$($scenario.name)-$axis"
                for ($index=0; $index -lt $scenario.starts.Count; $index++) {
                    $start = $scenario.starts[$index]
                    $end = $scenario.ends[$index]
                    if ($axis -eq 'horizontal') { $page.DrawRectangle($start,2,$end,3) | Out-Null }
                    else { $page.DrawRectangle(2,$start,3,$end) | Out-Null }
                }
                $cases += [ordered]@{ pageId=[string]$page.ID; name=$page.Name; axis=$axis; selection=$scenario.order; starts=$scenario.starts; ends=$scenario.ends }
            }
        }
        $source = @(Read-Positions $document)
        $document.SaveAs((Join-Path $directory 'original.vsdx')) | Out-Null
        foreach ($case in $cases) {
            $page = $document.Pages.ItemFromID([int]$case.pageId)
            $selection = $page.CreateSelection(0)
            foreach ($id in $case.selection) { $selection.Select($page.Shapes.ItemFromID([int]$id),2) }
            $selection.Distribute($(if ($case.axis -eq 'horizontal') { 0 } else { 4 }),$false)
        }
        $native = @(Read-Positions $document)
        $document.SaveAs((Join-Path $directory 'native.vsdx')) | Out-Null
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; cases=$cases; source=$source; native=$native } |
            ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
} finally {
    if ($document) { $document.Saved = $true; $document.Close() }
    $application.Quit()
}
Write-Output $directory
