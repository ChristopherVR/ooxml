// Framework-neutral, DOM-free pagination engine for DocumentModel. Approximates Word pagination; it does not claim Word layout parity.
export * from './units';
export * from './measure';
export { ligatureCss } from './ligatures';
export * from './text-breaks';
export * from './input';
export * from './result';
export * from './page-fields';
export * from './page-style-refs';
export { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout';
export { adjustForWidowOrphan, suppressesSpacing, widowControlEnabled } from './keep-rules';
export { layoutRow, splitRowAtHeight, type RowLayout, type RowSplit } from './table-layout';
export { PageCursor } from './page-cursor';
export { placeParagraph, type ParagraphPlacement } from './flow-paragraph';
export { placeTable } from './flow-table';
export { layoutSections } from './page-flow';
export { adaptDocumentModel } from './adapter';
export { border as layoutBorder } from './adapt-table';
export { paragraphFloats } from './adapt-floats';
export { shapePaint } from './shape-paint';
export {
	createTextBoxLayouts,
	layoutTextBoxChain,
	type TextBoxLayout,
	type TextBoxLayouts,
} from './text-box-layout';
export { layoutDocument, layoutDocumentModel, storyTextBoxLayouts } from './layout';
export { reviewSourceLength, type LayoutReviewOptions } from './review-display';
export { floatPosition, floatOrigin, FLOAT_WRAP_NOTE } from './floats';
export { positionFloats } from './position-floats';
export { cssFontStack } from './fonts';
export { collectDocxWebfontRuns } from './webfont-runs';
export { placeTab, type TabPlacement } from './tab-stops';
export { createLayoutSession, type LayoutSessionOptions } from './layout-session';
