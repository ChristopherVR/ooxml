# ooxml-core

Shared modern OOXML logic for the Office viewers: units, colour, geometry, packaging (OPC), the common XML model, DrawingML, charts, diagrams (SmartArt), math and encryption primitives.

- Legacy binary formats and the compound-file container live in [`ole2`](../ole2), not here.
- Each Office type has its own repository and model (`docx-viewer`, `pptx-viewer`, `xlsx-viewer` later) and consumes these packages.
- The extraction plan, phases and decisions are in `docx-viewer/docs/ooxml-core-plan.md`.

## Packages

| Package                      | Status                  | Purpose                                                              |
| ---------------------------- | ----------------------- | -------------------------------------------------------------------- |
| `@christophervr/ooxml-units` | seeded from docx-viewer | Branded Emu/Twips/points types, EMU constants and conversions        |
| `@christophervr/ooxml-color` | seeded from pptx-viewer | Hex/RGB/HSL primitives, linear sRGB, OOXML percent and angle parsing |

Further packages (`color`, `geometry`, `opc`, `xml`, `drawingml`, `chart`, `diagram`, `math`, `crypto`, `schema`) are added phase by phase; see the plan.

## Working here

```
bun install
bun run typecheck
bun run test
```

See `AGENTS.md` for the working agreements and `PROVENANCE.md` for where each module came from.
