// Format-neutral DrawingML (ECMA-376 Part 1, 20.1): the colour, fill, line, text-body and geometry
// readers and writers over the shared `xml` DOM, the colour resolver, gradient helpers and the
// theme model. `diagram`, `chart`, `docx` and `xlsx` build on it; nothing here knows about slides,
// documents, sheets or diagrams.
export * from './types';
export {
	attributeReader,
	booleanAttribute,
	descendants,
	integerAttribute,
	stringAttribute,
} from './dom';
export * from './drawing-color';
export * from './color-transforms';
export * from './preset-colors';
export * from './drawing-color-css';
export * from './drawing-color-brightness';
export * from './ordered-color-transforms';
export * from './drawing-fill';
export * from './node-view';
export * from './gradient-read';
export * from './line-read';
export * from './drawing-text';
export * from './drawing-geometry';
export * from './drawing-shadow';
export * from './effect-model';
export * from './drawing-effects';
export * from './effect-read';
export * from './effect-dag';
export * from './write-line';
export * from './write-color';
export * from './patch-color';
export * from './write-fill';
export * from './write-theme';
export * from './gradient-presets';
export * from './gradient-geometry';
export * from './rect-path-gradient';
export * from './shape-path-gradient';
export * from './theme-model';
export * from './theme';
export * from './format-scheme';
export * from './theme-color';
export * from './text-model';
export * from './text-element-view';
export * from './text-paragraph-codec';
export * from './text-run-codec';
export * from './text-body-codec';
export {
	parseDrawingFont,
	parseDrawingHyperlink,
	parseDrawingRunProperties,
} from './read-run-properties';
export * from './read-paragraph-properties';
export * from './read-text-body';
export * from './write-run-properties';
export * from './write-paragraph-properties';
export * from './write-text-body';

export * from './path-data';
export * from './custom-geometry';
export * from './picture-crop';
export * from './picture-effects';
export * from './picture-fill';
export * from './image-content-type';
export * from './text-3d-model';
export * from './read-text-3d';
export * from './write-text-3d';
export * from './camera-presets';
export * from './picture-layout';
export * from './picture-standard-effects';
export * from './picture-transfer';
export * from './picture-artistic';

export { inspectRasterImage, RasterImageError, RASTER_IMAGE_LIMITS } from './raster-image';
export type { RasterImageInfo, RasterImageOptions, RasterImageErrorCode } from './raster-image';
export { decodePng, jpegComponents, type DecodedRaster } from './raster-decode';
export { encodePng } from './raster-encode';
export * from './picture-hsl-pixels';

export {
	capCompensatedDashArray,
	dashPatternUnits,
	dashPenWidth,
	isDashedStroke,
} from './dash-pattern';
export type { DashSegmentInput } from './dash-pattern';
export * from './html-text-paste';
export * from './word-art';
export * from './word-art-styles';
export * from './word-art-svg';
export * from './word-art-layout';
export { renderSvgPreview } from './svg-preview';
export * from './text-warp-arch';
export * from './text-warp-path';
export * from './model3d';
export { isSelfContainedGlb } from './glb-import';
export * from './shape-palette';

/** Loads the shared gallery data once the host opens document design tools. */
export const loadThemePresets = () => import('./theme-presets');
