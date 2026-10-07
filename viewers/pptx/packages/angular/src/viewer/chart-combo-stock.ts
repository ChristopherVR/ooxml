/**
 * Thin re-export shim → vendored `pptx-viewer-shared`.
 *
 * `buildComboViewModel` / `buildStockViewModel` were extracted to
 * `pptx-viewer-shared` (`render/chart-combo-stock.ts`). This shim keeps the
 * historical Angular import path for the chart engine and its tests.
 */
export * from 'ooxml-ui/pptx/render/chart-combo-stock';
