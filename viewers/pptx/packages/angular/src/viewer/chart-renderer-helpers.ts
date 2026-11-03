/**
 * Thin re-export shim → vendored `pptx-viewer-shared`.
 *
 * The entire SVG-primitive chart engine (value-model, palette, layout, bar /
 * line / pie / scatter / radar geometry, the `ChartViewModel` contract and the
 * `Svg*` primitive descriptors) was extracted to `pptx-viewer-shared`
 * (`render/chart-view-model.ts`) and is now consumed by every binding.
 *
 * This shim preserves the historical Angular import surface (re-exported from
 * `internals.ts`) for the colocated tests. Library code imports these names
 * from `ooxml-ui/pptx/render/chart-view-model` directly: ng-packagr's
 * declaration bundler refers to a name imported through an `export *` of a
 * package as `<module>_d_exports.Name` and never declares that namespace.
 *
 * It re-exports from the vendored module file (not the `../internal/shared`
 * barrel) because the engine's low-level palette/range helpers deliberately
 * share names with `chart-helpers.ts` and are therefore not flattened into the
 * shared barrel.
 */
export * from 'ooxml-ui/pptx/render/chart-view-model';
