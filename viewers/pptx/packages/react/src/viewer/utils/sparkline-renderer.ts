/**
 * Thin re-export shim → `pptx-viewer-shared`.
 *
 * The inline mini-chart (sparkline) SVG renderer was consolidated into
 * `pptx-viewer-shared` (`render/chart-sparkline.ts`). This shim preserves the
 * historical React import surface so consumers and colocated tests keep
 * importing the same names unchanged.
 */
export type { SparklineData } from 'ooxml-ui/pptx';
export { renderSparklineSvg } from 'ooxml-ui/pptx';
