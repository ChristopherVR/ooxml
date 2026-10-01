# Working agreements

- This repository owns **modern OOXML** (ECMA-376 / ISO 29500: WordprocessingML, PresentationML, SpreadsheetML and the shared DrawingML, charts, diagrams, math, OPC packaging) logic that more than one Office viewer needs. It is a sibling of `ole2`.
- `ole2` owns the legacy compound-file and binary formats (DOC, XLS, PPT, CFB, RC4/MD4). Never move modern OOXML into `ole2`, and never move binary codecs here. The encrypted-package container is CFB (ole2); the encryption primitives are modern OOXML (here).
- Each Office type keeps its own repository and document model (`docx-viewer`, `pptx-viewer`, later `xlsx-viewer`). They consume packages from here; they must not fork or copy what lives here. Where a viewer still has its own copy, the extraction plan (`docx-viewer/docs/ooxml-core-plan.md`) says when it is replaced.
- One shared XML model and one package structure for every consumer. Parsers and writers are written once against that model; type-specific code (slides, paragraphs, sheets) stays in the viewer repositories.
- Bun workspaces, one package per boundary, TypeScript strict (including `exactOptionalPropertyTypes` and `noUncheckedIndexedAccess`), Vitest. Publish ESM and CJS because consumers need both.
- Keep source modules under 300 lines where practical. Add regression tests for parsing, preservation and round-trip behaviour; move tests with the code they cover.
- Unsupported features must be reported honestly. Never claim Office parity or lossless export without evidence.
- Record extraction provenance (source repository, path, commit, what changed) in `PROVENANCE.md` for every module moved here.
- No commits or publishing unless requested. Never publish fixtures or stale output.
