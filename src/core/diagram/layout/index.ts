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

// The layout interpreter's `dgm:choose` walkers, over the ordered-XML tree (a `RawXmlView` adapts a
// model's raw slots), and the `dgm:if` evaluation and axis navigation they decide with.
export {
	algorithmParameters,
	groupedChildren,
	orderedXmlView,
	type RawXmlView,
} from './smartart-choose-xml';
export {
	orderedAttributes,
	parseIteratorAttributes,
	parseOrderedWhen,
	parseWhenAttributes,
	validateSmartArtControlFlow,
} from './smartart-layout-control-flow';
export { activeBranch, nestedChooseBranch } from './smartart-layout-interpreter-choose-branch';
export {
	chooseAlgorithm,
	chooseAlgorithmOfType,
	chooseAlgType,
} from './smartart-layout-interpreter-choose-algorithm';
export {
	structuralChooseAlgDepth,
	tunnelsPastOwnCompositeSlot,
	type StructuralChooseAlgResult,
} from './smartart-layout-interpreter-choose-depth';
export { evaluateWhen, type WhenContext } from './smartart-layout-interpreter-when';
export { resolvePresentationOf } from './smartart-layout-interpreter-presof-choose';
export { resolveAxisCount, resolveAxisNodes } from './smartart-layout-interpreter-axis-count';
