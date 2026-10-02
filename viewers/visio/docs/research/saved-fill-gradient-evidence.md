# Saved horizontal fill gradient evidence

Review date: 2026-10-02. This is a bounded cached ShapeSheet formatting slice, not full gradient support or native Microsoft Visio visual parity.

## Implemented contract

- Complete saved `FillGradientEnabled=1`, `FillGradientDir=0`, `RotateGradientWithShape=1`, and `UseGroupGradient=0` caches normalize into the existing shared linear SVG paint model. A nonzero legacy fill pattern does not suppress a complete explicitly enabled gradient. Transparent `FillPattern=0` still suppresses paint.
- Saved angles are internal radians. Support is deliberately limited to horizontal zero and pi directions modulo complete turns, with `1e-12` radian tolerance for decimal cache rounding. Zero is the positive x direction. Both horizontal directions are independent of the unresolved clockwise/counterclockwise question. Non-horizontal saved angles remain diagnosed; no DrawingML angle-sign assumption was applied.
- Stop position and transparency must be finite values in `[0,1]`, with nondecreasing positions. Equal positions are retained. Literal RGB, built-in palette and document palette colors are accepted; unresolved colors and incomplete stops retain foreground fallback diagnostics. The first ten active rows are used, consistent with the documented Fill Gradient section limit. Deleted rows are not painted.
- Per-stop transparency is normalized to stop opacity. The saved gradient does not also multiply `FillForegndTrans`, avoiding applying the same saved transparency twice.
- Explicit root selection substitutes only existing `Themed` stop cells at matching section/row indices. `QuickStyleFillMatrix=0` selects stop transparency and position; explicit root color selection additionally selects stop color. Literal caches remain authoritative. No absent rows or cells are invented; deleted or unresolved root entries are not used. Formula-error metadata is still reported.

## Primary sources

- [FillGradientDir](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/e19c498a-5277-4add-9953-8b85cb2af250): linear versus radial/rectangular modes.
- [FillGradientAngle](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/0ed91a89-2152-41c4-a263-6821e3fa77ff): zero follows the positive x-axis.
- [Fill Gradient section](https://learn.microsoft.com/en-us/office/client-developer/visio/fill-gradient-section): each row represents a stop; only the first ten rows are used.
- [Gradient stop cells](https://learn.microsoft.com/en-us/office/client-developer/visio/cell-element-fill-gradient-sectionvisio-xml): stop color, transparency and position semantics.
- [GradientStopPosition](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/64030657-ec1c-4bcd-adbb-701bb24004fa): normalized positions and monotonic order.
- [QuickStyleFillMatrix](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/25689058-b1e7-4d3c-a833-0a4c7180f5f2): explicit zero selects root fill formatting, including stop transparency and position.
- [RotateGradientWithShape](https://learn.microsoft.com/en-us/office/client-developer/visio/rotategradientwithshape-cell-gradient-properties-section): rotation behavior.

## Reproducible evidence and limits

Core regressions: `saved-fill-gradient.test.ts`, `theme-root-gradient.test.ts`, `saved-fill-gradient-corpus.test.ts`. Viewer regression: `render-saved-fill.test.ts`. Generated tests establish numeric, inheritance, diagnostic and SVG contracts only.

The optional corpus check reads Apache POI commit `732120980140d5ed64b482c470e0b625cdb1ab15`, `test-data/diagram/60973.vsdx`, SHA-256 `c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03`. Style 8 contains enabled linear stops and angle `4.7123889803847`; this is evidence of a remaining saved-angle gap, not support for its direction. Page ID 10 shapes 2-6 retain their unsupported saved-gradient diagnostics. No fixture bytes are redistributed.

Run `VISIO_THEME_CORPUS_DIR=/path/to/apache-poi/fixtures node_modules/.bin/vitest run src/visio/saved-fill-gradient*.test.ts src/visio/theme-root-gradient.test.ts` in the core checkout. Run `npm test -- src/render-saved-fill.test.ts` in the viewer after rebuilding and linking the core.

No positive real-corpus visual improvement or native-reference equivalence is claimed for this horizontal slice. The corpus's cached 270-degree gradients, radial/rectangular/path fills, group gradients, canvas-fixed gradients, arbitrary formulas, hatch/custom fills and incomplete theme stop mixing remain outside the supported contract.
