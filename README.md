# @christophervr/ooxml-core

All the logic behind the Office viewers, as one package. The viewers (`docx-viewer`, `pptx-viewer`, later `xlsx-viewer`) contain only their UI and consume this.

- Legacy binary formats and the compound-file container live in [`ole2`](../ole2), not here.
- The migration plan, phases and decisions are in `docx-viewer/docs/ooxml-core-plan.md`.

## Areas

Each area is a subpath import (`import { parseXml } from '@christophervr/ooxml-core/xml'`); the root entry groups them by namespace (`import { xml } from '@christophervr/ooxml-core'`).

| Area       | Status           | Purpose                                                                                                                                                                            |
| ---------- | ---------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `units`    | from docx-viewer | Branded Emu/Twips/points types, EMU constants and conversions                                                                                                                      |
| `color`    | from pptx-viewer | Hex/RGB/HSL primitives, linear sRGB, OOXML percent and angle parsing                                                                                                               |
| `geometry` | from pptx-viewer | Preset shapes, connection sites, clip paths, callouts, boolean shape ops                                                                                                           |
| `xml`      | from docx-viewer | The shared XML model: strict DOM parse/serialize, namespaces, helpers                                                                                                              |
| `opc`      | from docx-viewer | Relationships, content types, part paths, zip helpers, safe hyperlinks                                                                                                             |
| `docx`     | from docx-viewer | WordprocessingML model, parser, serializer, editing, validation (ECMA-376 XSD checks in tests) |
| `pptx`     | from pptx-viewer | PresentationML model, parser, serializer, editing, converter, CLI, signatures (subpaths `/pptx`, `/pptx/converter`, `/pptx/cli`, `/pptx/signature-node`; relaxed TS flags for now) |

Planned areas: `drawingml`, `chart`, `diagram`, `math`, `crypto`, `schema`, then the document area `xlsx` (models, parsers, serializers, editing, layout) and `collab` (Yjs and the sync protocol).

## Working here

```
bun install
bun run typecheck
bun run test
bun run build
```

See `AGENTS.md` for the working agreements and `PROVENANCE.md` for where each module came from.
