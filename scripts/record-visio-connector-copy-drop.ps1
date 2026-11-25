param(
    [Parameter(Mandatory = $true)][string]$OutputDirectory
)
# ShapeSheet observations from an owned, invisible Visio instance.
# The drawings include Microsoft stencil masters and are for local comparison only.
$ErrorActionPreference = 'Stop'
$out = (New-Item -ItemType Directory -Force -Path $OutputDirectory).FullName
$app = New-Object -ComObject Visio.InvisibleApp
$app.AlertResponse = 7

function Read-Shape($shape, [bool]$geometryFormulas = $false) {
    $cells = [ordered]@{}
    foreach ($name in @('BeginX','BeginY','EndX','EndY','PinX','PinY','Width','Height','LocPinX','LocPinY','Angle','FlipX','FlipY','TxtPinX','TxtPinY','ShapeRouteStyle','ConFixedCode','ConLineJumpCode','ConLineJumpStyle','BegTrigger','EndTrigger','ObjType')) {
        if ($shape.CellExistsU($name, 0)) {
            $cell = $shape.CellsU($name)
            $cells[$name] = [ordered]@{ value=[double]$cell.ResultIU; formula=$cell.FormulaU }
        }
    }
    $geometry = @(for ($section = 10; $section -lt 10 + $shape.GeometryCount; $section++) {
        [ordered]@{ rows=@(for ($i = 1; $i -lt $shape.RowCount($section); $i++) {
            $values = @(for ($j = 0; $j -lt $shape.RowsCellCount($section, $i); $j++) {
                [double]$shape.CellsSRC($section, $i, $j).ResultIU
            })
            $row = [ordered]@{ type=$shape.RowType($section, $i); values=$values }
            if ($geometryFormulas) {
                $row.formulas = @(for ($j = 0; $j -lt $values.Count; $j++) { $shape.CellsSRC($section, $i, $j).FormulaU })
            }
            $row
        }) }
    })
    [ordered]@{
        id=[string]$shape.ID; type=$shape.Type; cells=$cells; geometry=$geometry
        glue=$shape.Connects.Count; children=@($shape.Shapes | ForEach-Object { Read-Shape $_ })
    }
}

function Set-Ends($shape, $x1, $y1, $x2, $y2) {
    $shape.CellsU('BeginY').FormulaU = [string]$y1
    $shape.CellsU('EndY').FormulaU = [string]$y2
    $shape.CellsU('BeginX').FormulaU = [string]$x1
    $shape.CellsU('EndX').FormulaU = [string]$x2
}

try {
    $doc = $app.Documents.Add('')
    $stencil = $app.Documents.OpenEx('CONNEC_U.VSSX', 66)
    $dynamic = $stencil.Masters.ItemU('Dynamic connector')
    $page = $doc.Pages.Item(1)
    $page.NameU = 'Crossing'
    $page.PageSheet.CellsU('LineJumpCode').ResultIU = 1
    $horizontal = $page.Drop($dynamic, 4, 4)
    Set-Ends $horizontal 1 4 7 4
    $a = $page.DrawRectangle(3.5, 1, 4.5, 2)
    $b = $page.DrawRectangle(3.5, 3.75, 4.5, 3.8)
    $vertical = $page.Drop($dynamic, 4, 4)
    $vertical.CellsU('BeginX').GlueTo($a.CellsU('PinX'))
    $vertical.CellsU('EndX').GlueTo($b.CellsU('PinX'))
    $crossing = [ordered]@{ source=(Read-Shape $horizontal); otherBefore=(Read-Shape $vertical) }

    $moves = foreach ($axis in @('X','Y')) {
        $page = $doc.Pages.Add()
        $page.NameU = "Obstacle $axis"
        $free = $page.Drop($dynamic, 4, 4)
        if ($axis -eq 'Y') {
            Set-Ends $free 1.125 4 7 4
            $obstacle = $page.DrawRectangle(3.5, 5.5, 4.5, 6.5)
        } else {
            Set-Ends $free 4 1.125 4 7
            $obstacle = $page.DrawRectangle(5.5, 3.5, 6.5, 4.5)
        }
        # Seed a direct route by removing the intermediate jog from the starting fixture.
        $free.DeleteRow(10, 3)
        $free.DeleteRow(10, 2)
        $obstacle.CellsU('ObjType').ResultIU = 1
        [ordered]@{ axis=$axis; pageId=[string]$page.ID; before=(Read-Shape $free); blocker=(Read-Shape $obstacle) }
    }

    $groupMaster = $doc.Masters.Add()
    $groupMaster.NameU = 'Grouped line'
    $masterEdit = $groupMaster.Open()
    $line = $masterEdit.DrawLine(1, 2, 3, 2)
    [void]$line.ConvertToGroup()
    $group = $masterEdit.Shapes.Item(1)
    [void]$group.DrawRectangle(0.85, -0.15, 1.15, 0.15)
    $masterEdit.Close()
    $cases = @()
    foreach ($name in @('Directed line 1','Angled connector','Curve connect 1','Elbow 1','Grouped line')) {
        $page = $doc.Pages.Add()
        $page.NameU = $name
        $master = if ($name -eq 'Grouped line') { $groupMaster } else { $stencil.Masters.ItemU($name) }
        $shape = $page.Drop($master, 4, 5)
        $cases += [ordered]@{
            name=$name; pageId=[string]$page.ID; masterId=[string]$shape.Master.ID
            master=(Read-Shape $shape.Master.Shapes.Item(1)); drop=(Read-Shape $shape)
        }
    }
    [void]$doc.SaveAs((Join-Path $out 'before.vsdx'))
    $crossing.copy = Read-Shape ($horizontal.Duplicate())
    $crossing.otherAfter = Read-Shape $vertical
    $crossing.gluedCopy = Read-Shape ($vertical.Duplicate())
    foreach ($move in $moves) {
        $page = $doc.Pages.ItemFromID([int]$move.pageId)
        $free = $page.Shapes.ItemFromID([int]$move.before.id)
        $obstacle = $page.Shapes.ItemFromID([int]$move.blocker.id)
        $obstacle.CellsU('Pin' + $move.axis).ResultIU = 4
        $move.blocked = Read-Shape $free
        $obstacle.CellsU('Pin' + $move.axis).ResultIU = 6
        $move.cleared = Read-Shape $free
    }
    foreach ($case in $cases) {
        $page = $doc.Pages.ItemFromID([int]$case.pageId)
        $case.copy = Read-Shape ($page.Shapes.ItemFromID([int]$case.drop.id).Duplicate())
    }
    [void]$doc.SaveAs((Join-Path $out 'after.vsdx'))
    $json = [ordered]@{ application='Microsoft Visio'; version=$app.Version; dynamic=(Read-Shape $horizontal.Master.Shapes.Item(1) $true); crossing=$crossing; moves=$moves; masters=$cases } |
        ConvertTo-Json -Depth 20 -Compress
    [IO.File]::WriteAllText((Join-Path $out 'connector-copy-drop.json'), $json + [Environment]::NewLine)
} finally {
    foreach ($document in @($app.Documents)) { $document.Saved = $true }
    $app.Quit()
}
