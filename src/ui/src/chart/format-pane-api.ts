/**
 * The chart Format pane's neutral model, for products that build their own
 * source or apply its edits: descriptor types, the source and view contracts,
 * the model builder, the `ChartSpace` source and the field-to-intent mapping.
 *
 * @module chart/format-pane-api
 */
export type * from './format-pane-types';
export type * from './format-pane-view';
export {
	translateChartFormatText,
	hexForInput,
	type ChartFormatTranslate,
} from './format-pane-common';
export {
	chartFormatPaneTitle,
	chartFormatTargetLabel,
	chartFormatTargetOptions,
	type ChartFormatDescribe,
} from './format-pane-targets';
export {
	buildChartFormatModel,
	chartFormatSourceTabs,
	resolveChartFormatSourceTarget,
} from './format-pane-model';
export {
	chartFormatFieldIntent,
	defaultPaneGradient,
	drawingColorFromHex,
	drawingGradientFrom,
} from './format-pane-intent';
export { chartSpaceFormatSource, type ChartSpaceFormatSourceOptions } from './format-pane-space';
