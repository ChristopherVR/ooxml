# Regenerates the fixed Visio-authored reference corpus (VSDX plus Page.Export PNGs) with an owned
# invisible Visio instance. Output defaults to viewers/visio/docs/evidence/reference-corpus.
#   reference-corpus.vsdx  six pages: basic shapes, master instances, connectors (stencil connector
#                          glued to shapes), groups, a container, a themed page
#   reference-flat.vsdx    three pages with no stencil masters: basic shapes, groups, glued lines
# Run: pwsh scripts/record-visio-reference-corpus.ps1
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\viewers\visio\docs\evidence\reference-corpus')
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
$flat = $null
$stencils = @()
function Open-Stencil($app, $path) {
    $stencil = $app.Documents.OpenEx($path, 66)   # visOpenRO + visOpenHidden
    $script:stencils += $stencil
    $stencil
}
function Add-Page($doc, $list, $name) {
    $page = if ($list.Count -eq 0) { $doc.Pages.Item(1) } else { $doc.Pages.Add() }
    $page.Name = $name
    $page
}
function Add-Basic($page) {
    $r = $page.DrawRectangle(1,6,3.5,7.5); $r.NameU = 'Basic.Rect'; $r.Text = 'Rectangle'
    $r.CellsU('FillForegnd').FormulaU = 'RGB(200,220,255)'
    $e = $page.DrawOval(4.5,6,7,7.5); $e.NameU = 'Basic.Oval'; $e.Text = 'Oval'
    $e.CellsU('FillForegnd').FormulaU = 'RGB(255,230,180)'
    $e.CellsU('LineWeight').FormulaU = '3 pt'
    $l = $page.DrawLine(1,4,7,5); $l.NameU = 'Basic.Line'
    $l.CellsU('LineColor').FormulaU = 'RGB(180,0,0)'
    $t = $page.DrawRectangle(1,2,4,3); $t.NameU = 'Basic.Rotated'; $t.Text = 'Rotated 30 deg'
    $t.CellsU('Angle').FormulaU = '30 deg'
}
function Add-Groups($page) {
    $m1 = $page.DrawRectangle(1,5,2.5,6.5); $m1.Text = 'M1'
    $m2 = $page.DrawOval(3,5,4.5,6.5); $m2.Text = 'M2'
    $m3 = $page.DrawRectangle(2,3,4,4.5); $m3.Text = 'M3'
    $sel = $page.CreateSelection(0)
    foreach ($s in @($m1,$m2,$m3)) { $sel.Select($s, 2) }
    $group = $sel.Group()
    $group.NameU = 'Group.Main'
    $solo = $page.DrawRectangle(6,3,8,4.5); $solo.Text = 'Ungrouped'
}
function Write-Pages($pages, $manifestPages) {
    foreach ($page in $pages) {
        $png = ($page.Name + '.png')
        $page.Export((Join-Path $script:directory $png))
        $info = [ordered]@{ name = $page.Name; png = $png; shapes = @() }
        foreach ($shape in $page.Shapes) {
            $info.shapes += [ordered]@{ id = $shape.ID; name = $shape.NameU; type = $shape.Type; text = $shape.Text
                PinX = [double]$shape.CellsU('PinX').ResultIU; PinY = [double]$shape.CellsU('PinY').ResultIU
                Width = [double]$shape.CellsU('Width').ResultIU; Height = [double]$shape.CellsU('Height').ResultIU }
        }
        $manifestPages.Add($info)
    }
}
try {
    $application.AlertResponse = 7
    $document = $application.Documents.Add('')
    $pages = @()
    $themeApplied = $true
    $page = Add-Page $document $pages 'basic-shapes'; $pages += $page; Add-Basic $page
    # master instances
    $page = Add-Page $document $pages 'master-instances'; $pages += $page
    $flow = Open-Stencil $application 'BASFLO_U.VSSX'
    $x = 2
    foreach ($name in @('Process','Decision','Start/End')) {
        $shape = $page.Drop($flow.Masters.ItemU($name), $x, 6)
        $shape.Text = $name
        $x += 2.5
    }
    # connectors with glue
    $page = Add-Page $document $pages 'connectors'; $pages += $page
    $a = $page.DrawRectangle(1,5,3,6); $a.NameU = 'Conn.A'; $a.Text = 'A'
    $b = $page.DrawRectangle(5,3,7,4); $b.NameU = 'Conn.B'; $b.Text = 'B'
    $c = $page.DrawRectangle(5,6,7,7); $c.NameU = 'Conn.C'; $c.Text = 'C'
    $connectorMaster = $flow.Masters.ItemU('Dynamic connector')
    foreach ($pair in @(@($a,$b),@($a,$c))) {
        $conn = $page.Drop($connectorMaster, 0, 0)
        $conn.CellsU('BeginX').GlueTo($pair[0].CellsU('PinX'))
        $conn.CellsU('EndX').GlueTo($pair[1].CellsU('PinX'))
    }
    $page = Add-Page $document $pages 'groups'; $pages += $page; Add-Groups $page
    # container
    $page = Add-Page $document $pages 'containers'; $pages += $page
    $k1 = $page.DrawRectangle(2,5,3.5,6); $k1.Text = 'In 1'
    $k2 = $page.DrawRectangle(4.5,5,6,6); $k2.Text = 'In 2'
    $containerStencil = Open-Stencil $application ($application.GetBuiltInStencilFile(2, 0))
    $sel = $page.CreateSelection(0)
    foreach ($s in @($k1,$k2)) { $sel.Select($s, 2) }
    $container = $page.DropContainer($containerStencil.Masters.Item(1), $sel)
    $container.NameU = 'Container.Main'
    # themed page
    $page = Add-Page $document $pages 'themed'; $pages += $page
    try { $page.SetTheme('Office Theme') } catch { $themeApplied = $false }
    $q1 = $page.DrawRectangle(1,5,3,6.5); $q1.Text = 'Themed 1'
    $q2 = $page.DrawOval(4,5,6,6.5); $q2.Text = 'Themed 2'
    [void]$document.SaveAs((Join-Path $directory 'reference-corpus.vsdx'))
    # flat document: no masters, so core edit admission has nothing to refuse for lack of master formulas
    $flat = $application.Documents.Add('')
    $flatPages = @()
    $page = Add-Page $flat $flatPages 'flat-basic-shapes'; $flatPages += $page; Add-Basic $page
    $page = Add-Page $flat $flatPages 'flat-groups'; $flatPages += $page; Add-Groups $page
    $page = Add-Page $flat $flatPages 'flat-glued-lines'; $flatPages += $page
    $ga = $page.DrawRectangle(1,5,3,6); $ga.NameU = 'Glue.A'; $ga.Text = 'A'
    $gb = $page.DrawRectangle(5,3,7,4); $gb.NameU = 'Glue.B'; $gb.Text = 'B'
    foreach ($s in @($ga,$gb)) { [void]$s.AddRow(7,-1,0); $s.CellsU('Connections.X1').FormulaU = 'Width*0.5'; $s.CellsU('Connections.Y1').FormulaU = 'Height*0.5' }
    $gl = $page.DrawLine(2,5.5,6,3.5); $gl.NameU = 'Glue.Line'
    $gl.CellsU('BeginX').GlueTo($ga.CellsU('Connections.X1'))
    $gl.CellsU('EndX').GlueTo($gb.CellsU('Connections.X1'))
    [void]$flat.SaveAs((Join-Path $directory 'reference-flat.vsdx'))
    $manifest = [ordered]@{
        application = 'Microsoft Visio'; version = $application.Version; build = $application.Build
        generated = (Get-Date -Format 'yyyy-MM-dd'); themeApplied = $themeApplied
        documents = $null
    }
    $list = New-Object System.Collections.Generic.List[object]
    Write-Pages $pages $list
    $docs = [ordered]@{}; $docs.Add('reference-corpus.vsdx', @($list.ToArray()))
    $list = New-Object System.Collections.Generic.List[object]
    Write-Pages $flatPages $list
    $docs.Add('reference-flat.vsdx', @($list.ToArray())); $manifest.documents = $docs
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $directory 'manifest.json') -Encoding utf8
    Write-Output $directory
} finally {
    foreach ($d in @($document, $flat)) { if ($null -ne $d) { try { $d.Saved = $true; $d.Close() } catch { } } }
    foreach ($s in $stencils) { try { $s.Close() } catch { } }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
