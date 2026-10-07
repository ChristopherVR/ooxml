# Owned invisible Visio oracle: Delete(0), background clearing and last-page replacement.
param([string]$OutputDirectory = (Join-Path $env:TEMP 'visio-page-delete'), [switch]$VerifyCore)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutputDirectory | Out-Null
$app = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-State($doc) {
    $pages = @()
    foreach ($page in $doc.Pages) {
        $sheet = $page.PageSheet
        $background = $page.BackPage
        $pages += [ordered]@{ id=[string]$page.ID; name=$page.Name; background=[bool]$page.Background
            backPage= $(if ($background -is [string]) { $background } elseif ($background) { $background.Name } else { '' })
            width=$sheet.CellsU('PageWidth').ResultIU; height=$sheet.CellsU('PageHeight').ResultIU
            drawingScale=$sheet.CellsU('DrawingScale').ResultIU; pageScale=$sheet.CellsU('PageScale').ResultIU
            shapes=$page.Shapes.Count }
    }
    [ordered]@{ pageCount=$doc.Pages.Count; pages=$pages }
}
try {
    $app.AlertResponse = 6
    $cases = @()
    foreach ($name in @('referenced','background','last')) {
        $document = $app.Documents.Add('')
        $first = $document.Pages.Item(1)
        if ($name -eq 'referenced') {
            $first.DrawRectangle(1,1,3,2) | Out-Null
            $second = $document.Pages.Add()
            $shape = $second.DrawRectangle(1,1,3,2)
            $shape.AddNamedRow(242,'Width',0) | Out-Null
            $shape.CellsU('User.Width').FormulaU = 'Pages[Page-1]!ThePage!PageWidth'
            $link = $shape.AddHyperlink(); $link.SubAddress = 'Page-1'
            $target = $first
        } elseif ($name -eq 'background') {
            $target = $document.Pages.Add(); $target.Background = -1
            $first.BackPage = $target.Name
        } else {
            $first.PageSheet.CellsU('PageWidth').FormulaU = '10 in'
            $first.PageSheet.CellsU('DrawingScale').FormulaU = '2 in'
            $first.Name = 'Custom'
            $first.DrawRectangle(1,1,3,2) | Out-Null
            $target = $first
        }
        $document.SaveAs((Join-Path $OutputDirectory "$name-original.vsdx")) | Out-Null
        $deletedId = [string]$target.ID
        $target.Delete(0)
        $native = Read-State $document
        if ($name -eq 'referenced') {
            $native.widthCache = $shape.CellsU('User.Width').ResultIU
            $native.link = $link.SubAddress
        }
        $document.SaveAs((Join-Path $OutputDirectory "$name-native.vsdx")) | Out-Null
        $document.Saved=$true; $document.Close(); $document=$null
        $accepted = $null
        if ($VerifyCore) {
            $document = $app.Documents.OpenEx((Join-Path $OutputDirectory "core-$name.vsdx"),128)
            $accepted = Read-State $document
            if ($accepted.pageCount -ne $native.pageCount) { throw "Incorrect $name page count" }
            for ($i=0; $i -lt $native.pages.Count; $i++) {
                # Replacement page IDs are package-owned, not Visio's internal allocator IDs.
                foreach ($key in @('name','background','backPage','width','height','drawingScale','pageScale','shapes')) {
                    if ($accepted.pages[$i][$key] -ne $native.pages[$i][$key]) { throw "Incorrect $name $key" }
                }
                if ($name -ne 'last' -and $accepted.pages[$i].id -ne $native.pages[$i].id) { throw "Changed surviving page ID in $name" }
            }
            if ($name -eq 'referenced') {
                $shape=$document.Pages.Item(1).Shapes.Item(1)
                $accepted.widthCache=$shape.CellsU('User.Width').ResultIU
                $accepted.link=@($shape.Hyperlinks | ForEach-Object {$_.SubAddress})[0]
                if ($accepted.widthCache -ne $native.widthCache -or $accepted.link -ne $native.link) { throw 'Incorrect frozen reference caches' }
            }
            $document.SaveAs((Join-Path $OutputDirectory "core-$name-native-reopened.vsdx")) | Out-Null
            $document.Saved=$true; $document.Close(); $document=$null
        }
        $cases += [ordered]@{ name=$name; deletedId=$deletedId; native=$native; accepted=$accepted }
    }
    [ordered]@{ application='Microsoft Visio'; version=$app.Version; cases=$cases } | ConvertTo-Json -Depth 8 |
        Set-Content (Join-Path $OutputDirectory 'evidence.json') -Encoding utf8
} finally {
    if ($document) { $document.Saved=$true; $document.Close() }
    $app.Quit()
}
Write-Output $OutputDirectory
