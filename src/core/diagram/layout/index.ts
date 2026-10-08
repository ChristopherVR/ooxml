// Layout results and text-fit helpers shared by the engine and the layout interpreters.
export type * from './smartart-layout-types';
export { resolveFontTable } from './smartart-font-table';
export {
	SMARTART_LINE_SPACING_FACTOR,
	fitWrappedFontSize,
	measureTextWidth,
	wrappedLineCount,
	wrappedWidestLineWidth,
} from './smartart-text-wrap-fit';
