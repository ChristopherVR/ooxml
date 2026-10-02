# Adopting `docx/layout` and `docx/load` in docx-viewer

The pagination engine and the document loading facade now live in this package. docx-viewer's private workspace packages `@christophervr/docx-layout` (`packages/layout`), `@christophervr/docx-document` (`packages/document`) and `@christophervr/docx-legacy` (`packages/legacy`) are replaced by imports in the next wave. This page lists exactly what changes. docx-viewer was not modified by the extraction.

## New entry points

| Import                                  | Replaces                                                        | Contents                                                                                                                                                                                                                                                              |
| --------------------------------------- | --------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `ooxml-core/docx/layout` | `@christophervr/docx-layout`                                    | `layoutDocumentModel`, `layoutDocument`, `layoutSections`, `adaptDocumentModel`, the `Layout*` input/result types, `TextMeasurer`, `LayoutFontSpec`, `createFakeMeasurer`, `fontMetrics`, `cssFontStack`, `ligatureCss`, unit helpers (`twipsToPx`, `pxToTwips`, ...) |
| `ooxml-core/docx/load`   | `@christophervr/docx-document` and `@christophervr/docx-legacy` | `detectDocumentFormat`, `loadDocument`, `loadLegacyDoc`, `LegacyDocError`, `DocumentFormat`                                                                                                                                                                           |

Both ship ESM and CJS with declarations. `docx/load` inlines the `@christophervr/ole2` codecs at build time (`tsup.config.ts` `noExternal`, as the pptx bundle does); `scripts/package-smoke.mjs` fails if the packed output still imports ole2. Consumers add no ole2 dependency, and the main `docx` entry never pulls the legacy reader in.

## What stays in docx-viewer

`createCanvasMeasurer` was the only DOM code in the layout package (canvas and a hidden `<span>` for ligature/small-caps shaping). It is UI and was not moved. Move it, unchanged, into the viewer (for example `packages/web-component/src/canvas-measurer.ts`) and have it implement `TextMeasurer` and `LayoutFontSpec` imported from `docx/layout`. It uses `cssFontStack` and `ligatureCss`, both exported from `docx/layout`. Header and footer content is not laid out by the engine at all: pages report `sectionIndex` and `pageInSection`, and the viewer picks the default, first or even slot from `model.sections[n].headers/footers` and renders it.

## Mechanical steps

1. Bump the `ooxml-core` dependency to the release that contains these areas (rebuild `../ooxml-core` and `bun install --force` while using the `file:` dependency; update `OOXML_CORE_REF` in CI).
2. Rewrite imports across `packages/*` and `tests`:
   - `@christophervr/docx-layout` -> `ooxml-core/docx/layout`
   - `@christophervr/docx-document` -> `ooxml-core/docx/load`
   - `@christophervr/docx-legacy` -> `ooxml-core/docx/load`
   - Keep `@christophervr/docx-core` (re-export of `/docx`); `docx/layout` takes the same `DocumentModel`.
3. Delete `packages/layout`, `packages/document`, `packages/legacy`, their `tsconfig.json` path mappings (`@christophervr/docx-layout`, `-document`, `-legacy`), and the workspace/build/bundle entries that name them. Remove the `@christophervr/ole2` dependency from the viewer's own `package.json` if nothing else imports it; the shared ole2 repository is unaffected.
4. Drop the `workspace:*` dependencies on the three packages from the framework packages and from the web component; bundlers now resolve the core package.
5. Run the viewer's unit, package and browser contract suites. Layout unit tests moved here and run with `createFakeMeasurer`; keep the viewer's browser specs, which exercise the real canvas measurer.

## Behavioural notes

- Public layout and loading behaviour is unchanged. Only the ole2 version differs (core pins `0.3.0`, the viewer pinned `0.2.0`); re-run the legacy `.doc` browser specs after the bump.
- The engine is deterministic for a given model and measurer (covered by `regression.test.ts`) and does not mutate the model.

## Limits to keep stating in the UI and docs

The engine approximates Word pagination with injected font metrics. It does not claim Word layout parity: results depend on the measurer, fonts, kerning and hyphenation differ from Word, wrapped floating pictures are positioned approximately (`FLOAT_WRAP_NOTE` is reported in `approximations`), and headers/footers are not measured, so they do not shrink the body area. `LayoutResult.approximations` lists the known deviations for a given run and must be surfaced rather than hidden. Legacy `.doc` import extracts main-body paragraph text only (no tables, headers, images or most formatting) and saves only same-length paragraph text edits; unsupported edits throw `LegacyDocError`.
