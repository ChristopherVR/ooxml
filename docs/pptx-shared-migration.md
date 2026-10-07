# PowerPoint shared migration

The current main branch already provides `src/core` and `src/ui` workspaces,
shared controls and viewer compatibility aliases. New migrations extend those
packages rather than importing another copy of the viewer engine.

## Extracted APIs

| Source                           | Shared owner                                        | API                                                                                               |
| -------------------------------- | --------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| Chart inspector grid edit policy | `src/core/pptx/core/utils/chart-grid-operations.ts` | Six `chartGrid*` immutable data operations via `ooxml-core/pptx`                                  |
| Colour-picker swatch catalogue   | `src/ui/src/form/color-swatches.ts`                 | `OFFICE_COLOR_SWATCHES`, `OFFICE_COLOR_SWATCH_HEXES`, `OfficeColorSwatch` via `ooxml-ui/controls` |

Grid operations remain in the PowerPoint area because they operate on the
PowerPoint chart model, including point-formatting policy. The palette is UI
data, shared without changing any product's existing palette or labels.

## Viewer adoption after release

The viewer depends on published packages. Keep its existing grid and palette
modules until the new shared exports are released. The first grid migration
changed the facade too early; the follow-up restores compatibility with the
published version. Once the APIs are published:

1. Update the appropriate viewer dependency range.
2. Replace the grid implementation with aliased `chartGrid*` re-exports from
   `pptx-viewer-core`, preserving the viewer's six existing names.
3. Replace the palette with re-exports from `ooxml-ui/controls`.
4. Run the existing shared module tests, all binding checks and browser suites.

The retained source modules are temporary compatibility implementations, not
new owners. Provenance records the extracted modules and their baseline commit.
No framework caller needs to change when adoption lands.
