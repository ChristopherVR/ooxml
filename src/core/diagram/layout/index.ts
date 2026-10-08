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
export {
	constraintAttributes,
	formatXsdDouble,
	parseConstraintAttributes,
	parseRuleAttributes,
	parseXsdDouble,
	ruleAttributes,
	validateSmartArtConstraintRules,
	type DiagramAttributeValue,
} from './smartart-constraint-rules';

// The constraint solver: resolves `dgm:constr` references (`refType`/`refFor`) across the whole
// layout definition, choose-aware, for the layout interpreter.
export {
	buildConstraintIndex,
	EMPTY_CONSTRAINT_INDEX,
	resolveConstraint,
	type ConstraintIndex,
	type IndexedConstraint,
} from './smartart-constraint-solver';
export { resolveRatioConstraint } from './smartart-constraint-ratio-fallback';
export { selectConstraints } from './smartart-constraint-branch-index';
export {
	clampByRules,
	findConstraint,
	ratioConstraint,
} from './smartart-layout-interpreter-constraints';

// Arrangement discovery: which `dgm:alg` drives a layout definition (hierarchy, a decidable
// choose, a slot-mapping composite, the first structural algorithm), and the iteration helpers.
export {
	discoverArrangement,
	resolveFlowDirection,
	type ArrangementKind,
	type ArrangementPlan,
	type FlowDirection,
} from './smartart-layout-interpreter-model';
export { selectArrangedNodes } from './smartart-layout-interpreter-flow';

// The node forest, item font sizing and the styled node builders every arranger shares.
export { buildForest, buildTree, treeDepth, treeWidth, type TreeNode } from './smartart-tree';
export { resolveHierarchyItemFontSizePx } from './smartart-layout-interpreter-hierarchy-fontfit';
export {
	resolveRoleFontSize,
	resolveSharedItemFontSize,
	type FontFitItem,
} from './smartart-layout-item-font-size';
export {
	resolveTieredItemFontSize,
	SMARTART_DESCENDANT_FONT_SCALE,
} from './smartart-layout-item-font-tier';
export { presetBoxNode, type PresetBoxNodeParams } from './smartart-layout-interpreter-preset-node';
export {
	circleNode,
	polygonNode,
	rectNode,
	styleContext,
	type StyleContext,
} from './smartart-layout-interpreter-render';
export { presetPolygonPoints } from './smartart-layout-shape-polygon';
