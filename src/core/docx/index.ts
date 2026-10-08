// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export * from './model';
export * from './highlight';
export { LIGATURE_VALUES, isLigatures, type Ligatures } from './ligatures';
export * from './language';
export {
	parseParagraphStyleCatalog,
	resolveParagraphFormatting,
	resolveStyleNumbering,
} from './paragraph-styles';
export * from './underline';
export { parseRunStyleCatalog } from './character-styles';
export { parseTableStyleCatalog } from './table-styles';
export { resolveTableStyleFormatting } from './resolve-table';
export { resolveRunFormatting, type RunFormattingContext } from './run-formatting';
export { parseTheme, parseColorSchemeMapping } from './theme';
export {
	resolveThemeColorToken,
	resolveThemeColorReference,
	applyThemeTint,
	applyThemeShade,
	isThemeColorToken,
} from './theme-color';
export { loadDocx } from './parse';
export {
	DIAGRAM_GRAPHIC_URI,
	type DocxDiagram,
	type DocxDiagramNode,
	type DocxDiagramPart,
	type DocxDiagramRendering,
} from './diagram';
export { diagramsIn } from './diagram-document';
export { CHART_GRAPHIC_URI, CHART_NOTICE, chartsIn, type DocxChart } from './chart';
export { diagramColorTheme, resolveDiagramColor } from './diagram-theme';
export { saveDocx } from './save';
export {
	assertValidDocumentModel,
	DocxModelValidationError,
	validateDocumentModel,
	type ValidationIssue,
} from './validate-model';
export * from './numbering-model';
export { parseNumberingCatalog, resolveNumberingLevel } from './numbering-parse';
export {
	computeListLabels,
	displayListLabel,
	formatListNumber,
	resolveParagraphNumbering,
} from './numbering-format';
export { headingListLevels, isHeadingListKind, type HeadingListKind } from './heading-list-kinds';
export { setListStartOverride } from './numbering-start-edit';
export {
	ensureListDefinition,
	createListDefinition,
	linkStylesToList,
	type ListKind,
} from './numbering-editing';
export { formatNoteNumber, numberNotesInOrder } from './notes';
export * from './review-formatting-display';
export * from './review-document-formatting';
export {
	acceptAllRevisions,
	acceptRevision,
	findRevision,
	listRevisions,
	rejectAllRevisions,
	rejectRevision,
	type RevisionEntry,
} from './revision-commands';
export { DEFAULT_TABLE_BORDERS } from './table-defaults';
export { fieldName } from './field-runs';
export { isFieldLocked } from './field-lock';
export { dateFieldResult, datePicture, formatWordDate } from './field-date';
export * from './table-of-contents';
export { resolveCellVisuals, type CellPlacement } from './table-visuals';
export { DEFAULT_STYLES_XML, defaultStyleCatalogs } from './default-styles';
export * from './generated/wml-simple-types';
export * from './units';
export * from './simple-types';
export { alignFromJustification, parseJustification } from './paragraph-alignment';
export type { ParagraphAlign, ParsedJustification } from './paragraph-alignment';
export { withParseWarnings, enumValue, reportParseWarning } from './parse-diagnostics';
export { PROPERTY_FIELDS, type DocumentProperties } from './core-properties';
export {
	captionParagraphs,
	isCaptionOf,
	tableOfFiguresInstruction,
	tocCaptionLabel,
} from './table-of-figures';
export * from './document-stats';
export * from './column-settings';
export * from './line-spacing';
export * from './section-layout';
export * from './attr-units';
export * from './case-transform';
export * from './page-setup-model';
export * from './page-size';
export * from './section-edit';
