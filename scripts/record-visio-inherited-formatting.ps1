# Native acceptance of explicit Inh XML variants derived from a native-authored drawing.
param(
    [string]$OutputDirectory = (Join-Path $env:TEMP ('visio-inherited-formatting-' + [guid]::NewGuid().ToString('N'))),
    [string]$CoreOutputPath
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$directory = (Resolve-Path -LiteralPath $OutputDirectory).Path
$application = New-Object -ComObject Visio.InvisibleApp
$document = $null
function Read-Styles($doc) {
    @($doc.Pages.Item(1).Shapes | ForEach-Object {
        $shape = $_
        [ordered]@{ id=[string]$shape.ID; text=$shape.Text; rows=@(for ($index=0;$index -lt $shape.RowCount(3);$index++) {
            [ordered]@{ style=[double]$shape.CellsSRC(3,$index,2).ResultIU; size=[double]$shape.CellsSRC(3,$index,7).ResultIU }
        }); align=[double]$shape.CellsU('Para.HorzAlign').ResultIU }
    })
}
try {
    $application.AlertResponse = 7
    if ($CoreOutputPath) {
        $reference = Get-Content -LiteralPath (Join-Path $directory 'evidence.json') -Raw | ConvertFrom-Json
        $document = $application.Documents.OpenEx((Resolve-Path -LiteralPath $CoreOutputPath).Path,202)
        $actual = @(Read-Styles $document)
        if ($actual.Count -ne $reference.native.Count) { throw 'Reopened shape count differs.' }
        for ($index=0;$index -lt $actual.Count;$index++) {
            $observed=$actual[$index]; $expected=$reference.native[$index]
            if ($observed.id -cne $expected.id -or $observed.text -cne $expected.text -or
                $observed.align -ne $expected.align -or $observed.rows.Count -ne $expected.rows.Count) {
                throw "Native inherited formatting mismatch: shape $index"
            }
            for ($row=0;$row -lt $observed.rows.Count;$row++) {
                if ($observed.rows[$row].style -ne $expected.rows[$row].style -or
                    [Math]::Abs($observed.rows[$row].size - $expected.rows[$row].size) -gt 1e-10) {
                    throw "Native inherited formatting mismatch: shape $index row $row"
                }
            }
        }
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; accepted=$actual } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'reopen-evidence.json') -Encoding utf8
    } else {
        $document = $application.Documents.Add('')
        $style = $document.Styles.Add('Inherited Emphasis','',1,0,0)
        $style.CellsU('Char.Style').FormulaU='2'
        $style.CellsU('Para.HorzAlign').FormulaU='2'
        foreach ($index in @(1,2,3)) {
            $shape = $document.Pages.Item(1).DrawRectangle(1,$index*2,5,$index*2+1)
            $shape.Text='Inherited text'
            $shape.TextStyle=$style.Name
            $shape.CellsU('Char.Size').FormulaU='18 pt'
            if ($index -eq 3) { $range=$shape.Characters; $range.Begin=10; $range.End=14; $range.CharProps(7)=15 }
        }
        [void]$document.SaveAs((Join-Path $directory 'native-base.vsdx'))
        $document.Close(); $document=$null
        Copy-Item -LiteralPath (Join-Path $directory 'native-base.vsdx') -Destination (Join-Path $directory 'original.vsdx')
        # These documented XML variants are deliberate fixture edits, not native save output.
        $zip=[IO.Compression.ZipFile]::Open((Join-Path $directory 'original.vsdx'),[IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry=$zip.GetEntry('visio/pages/page1.xml')
            $reader=New-Object IO.StreamReader($entry.Open())
            $xml=New-Object Xml.XmlDocument
            $xml.PreserveWhitespace=$true
            $xml.LoadXml($reader.ReadToEnd()); $reader.Dispose()
            foreach ($shape in $xml.SelectNodes('//*[local-name()="Shape"]')) {
                foreach ($sectionName in @('Character','Paragraph')) {
                    $section=$shape.SelectSingleNode("*[local-name()='Section' and @N='$sectionName']")
                    if ($null -eq $section) { $section=$xml.CreateElement('Section',$xml.DocumentElement.NamespaceURI);$section.SetAttribute('N',$sectionName);[void]$shape.InsertBefore($section,$shape.SelectSingleNode("*[local-name()='Text']")) }
                    if ($section.ChildNodes.Count -eq 0) { $row=$xml.CreateElement('Row',$xml.DocumentElement.NamespaceURI);$row.SetAttribute('IX','0');[void]$section.AppendChild($row) }
                    foreach ($row in $section.SelectNodes("*[local-name()='Row']")) {
                        $cellName=if($sectionName -eq 'Character'){'Style'}else{'HorzAlign'}
                        foreach ($existing in @($row.SelectNodes("*[local-name()='Cell' and @N='$cellName']"))) { [void]$row.RemoveChild($existing) }
                        $cell=$xml.CreateElement('Cell',$xml.DocumentElement.NamespaceURI)
                        $cell.SetAttribute('N',$cellName)
                        $cell.SetAttribute('F','Inh')
                        if ($shape.GetAttribute('ID') -ne '2') { $cell.SetAttribute('V','2') }
                        [void]$row.AppendChild($cell)
                    }
                }
            }
            $entry.Delete(); $entry=$zip.CreateEntry('visio/pages/page1.xml')
            $writer=New-Object IO.StreamWriter($entry.Open(),[Text.UTF8Encoding]::new($false));$writer.Write($xml.OuterXml);$writer.Dispose()
        } finally { $zip.Dispose() }
        $document=$application.Documents.OpenEx((Join-Path $directory 'original.vsdx'),202)
        $source=@(Read-Styles $document)
        foreach ($shape in $document.Pages.Item(1).Shapes) {
            for ($row=0;$row -lt $shape.RowCount(3);$row++) { $shape.CellsSRC(3,$row,2).FormulaU='3' }
            $shape.CellsU('Para.HorzAlign').FormulaU='0'
        }
        [void]$document.SaveAs((Join-Path $directory 'native.vsdx'))
        [ordered]@{ application='Microsoft Visio'; version=$application.Version; sourceKind='native-derived Inh XML variants'; source=$source; native=@(Read-Styles $document) } |
            ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $directory 'evidence.json') -Encoding utf8
    }
    Write-Output $directory
} finally {
    if ($null -ne $document) { $document.Saved=$true;$document.Close() }
    $application.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
}
