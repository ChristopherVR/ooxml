# Whole-text spacing equal to the first paragraph

The source XML and observed PowerPoint paragraph properties are recorded in
`evidence.json`. Paragraph 1 directly sets 1.2; paragraph 2 inherits 1.5 from
the list style. A whole-text 1.2 command must change paragraph 2 even though
the displayed first-paragraph value already matches the command. PowerPoint
saved a direct `120000` percentage on paragraph 2 and retained the `150000`
list style. The core-exported file reopened with both paragraphs at 1.2.

Reproduce from the repository root after `bun install --frozen-lockfile` and
`bun run --cwd src/core build`. Generate into a new scratch directory:

```powershell
bun e2e/pptx/native-line-spacing/generate.ts C:/Temp/spacing-probe
```

With no existing PowerPoint process, run the following in PowerShell. Run
again for `core-edited.pptx` to measure the engine export before any native
edit. Use a new output name each time; retain the original generated files.

```powershell
if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Close PowerPoint first' }
$app = $null
$deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Open((Get-Item -LiteralPath 'C:/Temp/spacing-probe/source.pptx').FullName, 0, 0, 0)
    $shape = @($deck.Slides.Item(1).Shapes | Where-Object {
        $_.HasTextFrame -eq -1 -and $_.TextFrame.HasText -eq -1 -and
        $_.TextFrame.TextRange.Text.StartsWith('First')
    })[0]
    foreach ($n in 1, 2) { $shape.TextFrame.TextRange.Paragraphs($n, 1).ParagraphFormat | Select-Object SpaceWithin, LineRuleWithin }
    $shape.TextFrame.TextRange.ParagraphFormat.LineRuleWithin = -1
    $shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.2
    foreach ($n in 1, 2) { $shape.TextFrame.TextRange.Paragraphs($n, 1).ParagraphFormat | Select-Object SpaceWithin, LineRuleWithin }
    $deck.SaveAs('C:\Temp\spacing-probe\native-edited.pptx', 24)
} finally {
    if ($deck) { $deck.Close(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck) | Out-Null }
    if ($app) { $app.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
```

Inspect `ppt/slides/slide1.xml` in the native output for the paragraph and list
style XML recorded in `evidence.json`. This is a property and package check,
not a pixel or timing claim. The browser spec additionally verifies the
actual ribbon action, document Undo/Redo and exported original/edited XML
across five bindings. It uses the visible Undo button because Save can return
focus to a ribbon select, where document keyboard shortcuts ignore Undo.
Existing keyboard history tests remain unchanged.
