/** The DOM-free chart painter: a chart summary or a parsed chart part to a standalone SVG. */
export * from './summary';
export * from './summary-from-space';
export * from './summary-from-chartex';
export * from './render-space';
export * from './chart-view-markers';
export * from './chart-svg-markers';
export {
	chartSummaryView,
	categoryTotals,
	type ChartRefValue,
	type ChartSeriesView,
	type ChartViewModel,
	type ChartViewOptions,
	type EvaluateRef,
	type ValueAxisView,
} from './chart-view';
export type {
	ChartErrorBarsView,
	ChartStrokeView,
	ChartTrendlineView,
} from './chart-view-analysis';
export { renderChartSvg, type ChartSvgOptions } from './chart-svg';
export {
	chartAccessibleName,
	chartDataDescription,
	type ChartDataReadout,
} from './chart-accessibility';
export {
	chartPicturePaint,
	chartPicturePattern,
	chartPictureSize,
	type ChartPictureBox,
	type ChartPictureMode,
	type ChartPicturePaint,
	type ChartPictureSource,
} from './chart-picture-paint';
export { autoSeriesColor, chartColorScheme, modulateLuminance } from './chart-colors';
export { THEME_SLOTS, DEFAULT_THEME_COLORS, themeColor } from './theme-palette';
export {
	resolveChart3DProjection,
	type Chart3DViewInput,
	type Chart3DProjection,
	type Chart3DObliqueProjection,
	type Chart3DPerspectiveProjection,
} from './chart-3d-projection';
export { shadeChart3DFace, buildChart3DBoxFaceColors, type Chart3DFace } from './chart-3d-shading';
export { createChart3DCamera, type Chart3DCamera, type Chart3DSize } from './chart-3d-camera';
export {
	boxMesh,
	extrudeMesh,
	type Chart3DMesh,
	type Chart3DMeshFace,
	type Vec3,
} from './chart-3d-mesh';
export { chart3DApplies, chart3DSvg } from './chart-3d-svg';
export { barShapeMesh, type BarShapeBounds } from './chart-3d-shapes';
export { chartContourSvg } from './chart-3d-surface';
export type { Chart3DSurfaceStyle, Chart3DSurfacesView } from './chart-3d-walls';
