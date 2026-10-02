# Explicit root-style and solid theme-line evidence

Review date: 2026-10-02. This is a bounded saved-formatting slice implemented in `ooxml-core/visio`, not a formula evaluator or a native Visio visual-parity claim. The viewer consumes the normalized styles without reparsing theme XML.

## Implemented behavior and primary sources

- Saved `ColorSchemeIndex=0` selects root colors/formats. Saved `65534` uses the page's selector, including explicitly assigned PageSheet styles. Root identity is verified as style ID `0`, NameU `No Style`. Missing or malformed selectors do not become implicit root selection. [ColorSchemeIndex](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/1e7e9b7e-d116-41c0-9c53-5e57c26042a4), [ConnectorSchemeIndex root-format rule](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/9753d977-a1de-49e8-888c-cf532efe7982), [root style definition](https://learn.microsoft.com/en-au/openspecs/sharepoint_protocols/ms-vsdx/f1fbf678-12fb-40d4-b793-aece1c7881a9), [PageSheet style inheritance](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/f81673b1-da84-4754-b19e-a0475d889120).
- Saved `QuickStyleLineMatrix=0` independently selects supported root line properties. Replacement is limited to explicit `Themed` values: saved colors, width, cap, pattern, line transparency, and supported arrow/size caches. Existing literal/numeric results remain authoritative. Missing caches remain missing. [QuickStyleLineMatrix](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/edfd9f33-fea3-4bf5-8cd5-a91a3d677b03).
- Explicit root selection cannot fall through into a populated dynamic theme when its root cache is absent, invalid, or itself `Themed`. Formula-error metadata remains diagnosed when the last valid root value is used. Root Character lookup is cached per immutable parsed root and respects deleted sections/rows.
- A selected DrawingML line with exactly one explicit `prstDash val="solid"` and no competing custom dash resolves as solid. Missing, ambiguous, malformed and non-solid dash records remain diagnosed. No new preset dash lengths were inferred. Existing cached Visio patterns 2-23 retain their compatibility-spacing diagnostic. [DrawingML preset dash](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetdash), [Visio LinePattern](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-vsdx/f718c500-6066-4fcc-8b66-4bb5015f133f).

## Reproducible evidence

Core files: `theme-root.ts`, `theme-resolve.ts`, `theme-line.ts`; regression files: `theme-root.test.ts`, `theme-root-fallthrough.test.ts`, `theme-root-corpus.test.ts`, `theme-line.test.ts`, and `line-pattern.test.ts`.

The optional external tests use Apache POI commit `732120980140d5ed64b482c470e0b625cdb1ab15`, [test-data/diagram](https://github.com/apache/poi/tree/732120980140d5ed64b482c470e0b625cdb1ab15/test-data/diagram). No fixture bytes are committed.

- `60489.vsdx`, SHA-256 `15494702ecb5554c1dfaee2c09ac35e2ea656e2eed7ba4ba0cf1039e6125fbe1`: shape 16 resolves the root width `0.01041666666666667`, round cap and solid pattern. Shape 2 retains its literal width `0.006944444444444444`; group 111 retains its independently selected theme fill `#5b9bd5` and its unresolved text-color diagnostic.
- `60973.vsdx`, SHA-256 `c61ca252ea251262f81b18fb0e461c50797bf4b148b2c01448792447ada51f03`: shape 11 resolves the same saved root width, round cap and solid pattern. Unrelated formula, media and rendering limitations remain diagnosed.

From the core checkout, run `VISIO_THEME_CORPUS_DIR=/path/to/apache-poi/fixtures node_modules/.bin/vitest run src/visio`. The final local run passed 967 tests, including both hash-pinned root corpus tests; two separate opt-in EMF tests were skipped. Strict TypeScript and scoped formatting passed. Thirty independent root-selection, dash and error-propagation review checks also passed.

A 19-file corpus diagnostic pass reported eight unresolved colors and no unresolved caps/patterns after this slice. This is diagnostic coverage, not visual equivalence. The earlier 60973 pass hit the diagnostic limit, so total warning counts before/after are not directly comparable.

## Remaining limits

No group-theme inheritance rule was inferred from the corpus. No arbitrary formulas, missing numeric caches, non-solid theme dash spacing, custom master patterns, complete legacy/dynamic theme variants, theme fonts, or effects were added. Existing effects and other unsupported content remain diagnosed. Native Visio reference renders and browser visual comparisons were not performed for this slice. Generated regressions and exact normalized values establish the stated contracts only.
