# Provenance

Every module moved into this repository is recorded here: source repository and path, the commit it was taken at, and what was changed on the way in. Consumers replace their own copy only after the moved module's tests pass here.

| Package                      | Module                      | Source                                   | Source commit                | Changes                                                                                                                                             |
| ---------------------------- | --------------------------- | ---------------------------------------- | ---------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `@christophervr/ooxml-units` | `units.ts`, `units.test.ts` | `docx-viewer/packages/core/src/units.ts` | `c56b3ae` (main, 2026-10-01) | Verbatim copy; header comment reworded; `EMU_PER_*` constants added (previously duplicated in `docx-core/drawing.ts` and pptx `core/constants.ts`). |
