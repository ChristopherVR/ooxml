// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
export { mergeDocxSnapshots } from './merge-snapshots';
// The mail-merge engine is the `ooxml-core/docx/mail-merge` subpath; only its settings live here.
export type {
	MailMergeFieldMapping,
	MailMergeSettings,
	MailMergeSourceReference,
} from './mail-merge/settings';
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
export { resolveTableCellMargins } from './table-margins';
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
export { DEFAULT_DOCX_LOAD_LIMITS, type DocxLoadLimits } from './package-limits';
export { cellSourceBlocks } from './cell-blocks';
export {
	DIAGRAM_GRAPHIC_URI,
	type DocxDiagram,
	type DocxDiagramNode,
	type DocxDiagramPart,
	type DocxDiagramRendering,
} from './diagram';
export { diagramsIn } from './diagram-document';
export {
	changeDocxDiagramLayout,
	formatDocxDiagram,
	createDocxDiagram,
	editDocxDiagram,
	insertedDiagramPartNames,
	withDiagramNodes,
	type DocxDiagramCreation,
	type DocxDiagramEdit,
	type DocxDiagramEditResult,
} from './diagram-edit';
export { treeOrderDiagramNodes } from './diagram-nodes';
export {
	CHART_GRAPHIC_URI,
	CHART_NOTICE,
	chartsIn,
	documentDrawingBlocks,
	type DocxChart,
} from './chart';
export { docxChartPaint, themeCatalogColorScheme, type DocxChartPaint } from './chart-paint';
export {
	applyDocxChartDataSheetEdits,
	docxChartDataSheet,
	editDocxChart,
	isDocxChartEditable,
	isNewDocxChart,
	newChartImage,
	newDocxChart,
	NEW_CHART_SIZE_PX,
} from './chart-edit';
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
export { pictureBulletSizePt } from './numbering-picture';
export {
	DRAWING_WRAPS,
	drawingWrapOf,
	placementForWrap,
	setDrawingWrap,
	type DrawingWrap,
	type FloatOrigin,
} from './drawing-wrap';
export {
	alignDrawingBox,
	alignDrawingToPage,
	distributeDrawingBoxes,
	drawingBounds,
	drawingMovement,
	drawingPosition,
	drawingStackOrder,
	groupDrawingShapes,
	placeDrawingBox,
	rotateDrawing,
	ungroupDrawingShapes,
	type DrawingAlignment,
	type DrawingAlignTarget,
	type DrawingBox,
	type DrawingPosition,
	type DrawingTextDirection,
	type DrawingTransform,
} from './drawing-arrange';
export {
	isValidPictureCrop,
	pictureCropOf,
	setPictureCrop,
	uncroppedPictureSize,
} from './picture-crop-edit';
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
	createPictureBulletList,
	linkStylesToList,
	type ListKind,
} from './numbering-editing';
export { formatNoteNumber, numberNotesInOrder } from './notes';
export { PICTURE_WATERMARK_SCALES, pictureWatermarkSizePt } from './picture-watermark';
export {
	convertNoteRun,
	convertNotes,
	planNoteConversion,
	retargetNoteReference,
	type NoteConversion,
	type NoteConversionPlan,
	type NoteSelection,
} from './note-convert';
export * from './review-formatting-display';
export * from './review-document-formatting';
export * from './format-change-description';
export { watermarkSizePt } from './watermark';
export {
	autoCorrectTextInput,
	type AutoCorrectAction,
	type AutoCorrectOptions,
} from './autocorrect';
export {
	acceptAllRevisions,
	acceptRevision,
	findRevision,
	listRevisions,
	rejectAllRevisions,
	rejectRevision,
	type RevisionEntry,
} from './revision-commands';
export {
	compareDocuments,
	UNCOMPARED_DOCUMENT_PARTS,
	revisionSpans,
	type CompareDocumentsOptions,
	type CompareDocumentsResult,
	type CompareGranularity,
	type CompareSummary,
	type RevisionSpan,
	type UncomparedDocumentPart,
} from './compare/index';
export { DEFAULT_TABLE_BORDERS } from './table-defaults';
export {
	setTableAlignment,
	setTableBorders,
	setTableGridWidths,
	setTableIndent,
	setTableLook,
	setTableStyle,
	setTableWidth,
	updateTableProperties,
	type TableBordersPatch,
	type TableLookPatch,
	type TablePropertiesPatch,
	type TableWidthSpec,
} from './table-properties-edit';
export { fieldName } from './field-runs';
export { isFieldLocked } from './field-lock';
export { dateFieldResult, datePicture, formatWordDate } from './field-date';
export {
	parseStyleReference,
	styleRefSource,
	type StyleRefSource,
	type StyleReference,
} from './style-ref';
export { fieldResultRun, withFieldCodeFormatting } from './field-result-spans';
export { bodyStyleRefResult, refreshStyleRefFields } from './style-ref-fields';
export {
	COMPUTED_FIELD_NAMES,
	computedFields,
	refreshComputedFields,
	type ComputedFieldResult,
	type ComputedFields,
} from './computed-fields';
export { docPropertyResult, fieldProperties } from './doc-property-field';
export {
	listFormFields,
	type FormFieldChoice,
	type FormFieldInfo,
	type FormFieldStory,
	type FormFieldType,
} from './form-fill-fields';
export { fillFormFields, type FormFillOptions, type FormFillResult } from './form-fill-write';
export type { FormFieldValue } from './form-fill-values';
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
export {
	INSPECTION_CATEGORIES,
	inspectDocument,
	type CategoryInspection,
	type InspectionCategory,
	type InspectionFinding,
	type InspectionFindingKind,
} from './inspect-document';
export { removeDocumentInformation } from './remove-document-information';
export type { PackageInformation } from './package-information';
export * from './column-settings';
export {
	createProtectionPassword,
	enforceProtection,
	verifyProtectionPassword,
} from './protection';
export * from './line-spacing';
export * from './section-layout';
export * from './attr-units';
export * from './case-transform';
export * from './page-setup-model';
export * from './page-size';
export * from './section-edit';
export * from './content-control-edit';
export {
	formCheckboxChecked,
	formDropDownIndex,
	formFieldDisplayText,
	formFieldOfRun,
	selectFormDropDownEntry,
	setFormCheckbox,
} from './form-fields';
export { parseContentControl } from './content-control';
export * from './shape-model';
export * from '../drawingml/shape-palette';
export {
	drawingShapeSvg,
	shapeSvg,
	shapeAppearance,
	shapeColorCss,
	shapeTextColor,
} from './shape-svg';
export { simpleTextBoxDrawing } from './text-box-drawing';

export * from './chart-ex-edit';
export * from './clipboard/index';
export {
	applyParagraphFormatPatch,
	applyRunFormatPatch,
	hasFormatCriteria,
	hasReplaceFormat,
	matchesFormatCriteria,
	matchesParagraphFormat,
	normalizeFormatColor,
	type FindFormatContext,
	type FindFormatCriteria,
	type ReplaceFormatPatch,
} from './find-format-criteria';
export {
	findDocumentFormattedMatches,
	findFormattedMatches,
	flattenBlocksForFind,
	modelFormatContext,
	type FindFormatOptions,
	type FormattedFindMatch,
} from './find-format';
export {
	replaceFormattedMatches,
	type ReplaceFormatOptions,
	type ReplaceFormatResult,
} from './find-format-replace';
export {
	replaceDocumentFormattedMatches,
	type DocumentReplaceFormatResult,
} from './find-format-document';

export type { LatentStyleException, LatentStyles } from './latent-styles';

export {
	canBaseStyleOn,
	editStyleCatalogs,
	paragraphStyleId,
	styleCatalogsOf,
	stylesOf,
	type StyleCatalogs,
	type StyleEdit,
	type StyleKind,
} from './edit-styles';
export * from './document-design';
export * from './design-defaults';
