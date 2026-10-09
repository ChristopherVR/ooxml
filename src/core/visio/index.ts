/** VSDX parsing/editing and conservative legacy VSD preview scenes. */
export type * from './model';
export type { VisioShapeData, VisioHyperlink, VisioMetadataOptions } from './shape-metadata';
export { parseVsdx, getVisioPageLayers, type ParseVsdxOptions } from './parser';
export { loadVisio } from './load';
export { parseLegacyVsd, type ParseLegacyVsdOptions } from './legacy';
export { VisioPackageError, type VisioPackageLimits } from './package';
export {
	resolveVisioPageVisibility,
	VISIO_VISIBILITY_LIMITS,
	type VisioVisibilityOptions,
	type VisioResolvedShapeVisibility,
} from './visibility';
export {
	inspectVisioRasterImage,
	VisioImageError,
	VISIO_RASTER_IMAGE_LIMITS,
	type VisioImageOptions,
	type VisioImageErrorCode,
	type VisioRasterImageInfo,
} from './media';
export {
	inspectVisioEmfAdmission,
	VISIO_EMF_ADMISSION_LIMITS,
	type VisioEmfAdmissionOptions,
	type VisioEmfAdmissionResult,
	type VisioEmfAdmissionStatus,
	type VisioEmfAdmissionDiagnostic,
	type VisioEmfAdmissionMetrics,
} from './emf-admission';
export {
	sanitizeVisioForeignVectorTree,
	validateVisioForeignVector,
	VisioForeignVectorError,
	VISIO_FOREIGN_VECTOR_LIMITS,
	type VisioForeignVector,
	type VisioForeignVectorLimits,
	type VisioForeignVectorCommand,
	type VisioForeignVectorClip,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorNode,
	type VisioForeignVectorPath,
	type VisioForeignVectorPaint,
	type VisioForeignVectorMatrix,
} from './foreign-vector';
export {
	convertVisioMetafile,
	type VisioMetafileTreeConverter,
	type VisioMetafileConversionResult,
} from './convert-metafile';

export {
	editVsdx,
	type VisioEdit,
	type VisioGeometryEdit,
	type VisioTextEdit,
	type VisioTextRange,
	type VisioTextRangesEdit,
	type VisioPageInsert,
	type VisioPageReorder,
	type VisioPageRename,
	type VisioPageDelete,
	type VisioPageSizeEdit,
	type VisioPageSetupEdit,
	type VisioPagePropertiesEdit,
	type VisioPageDecorationEdit,
	type VisioPageSetupEdits,
	type VisioScaleUnit,
	type VisioPageEdit,
	type VisioFormatEdit,
	type VisioTextFormatEdit,
	type VisioShapeFormatEdit,
	type VisioShapeOrderEdit,
	type VisioDuplicateShapesEdit,
	type VisioPasteShapesEdit,
	type VisioMetadataEdit,
	type VisioPictureInsertEdit,
	type VisioShapeHyperlinkEdit,
	type VisioShapeScreenTipEdit,
	type VisioHyperlinkFields,
	type VisioGroupEdit,
	type VisioGroupShapesEdit,
	type VisioUngroupShapeEdit,
	type VisioResizeAnchor,
	type VisioChangeShapeEdit,
	type VisioChangeShapeTarget,
	type VisioConnectorGlue,
	type VisioPageThemeEdit,
	type EditVsdxOptions,
	type EditVsdxResult,
} from './edit';
export {
	captureVisioClipboard,
	serializeVisioClipboard,
	deserializeVisioClipboard,
	VISIO_CLIPBOARD_MAGIC,
	VISIO_CLIPBOARD_MAX_CHARS,
	type VisioClipboardSnapshot,
} from './clipboard';
export { createVsdx, type CreateVsdxOptions } from './create-document';
export { VISIO_PATH_SEGMENT_LIMIT, type VisioPathCreateEdit } from './edit-path-commands';
export {
	visioSimplifyPath,
	visioFitFreeform,
	visioFitPencil,
	visioQuarterArc,
	visioPathSamples,
	visioPathBounds,
	visioCircleThrough,
	type VisioPathPoint,
	type VisioPathSegment,
} from './path-fit';
export {
	VISIO_QUICK_STYLE_COLORS,
	VISIO_QUICK_STYLE_VARIANT_COLORS,
	VISIO_SHADOW_PRESETS,
	isVisioQuickStyleColor,
	isVisioShadowPreset,
	visioFallbackQuickStyle,
	visioShadowPresetGeometry,
	type VisioQuickStyle,
	type VisioQuickStyleColor,
	type VisioShadowPreset,
} from './edit-formatting-effects';
export {
	VISIO_GLOW_SIZES,
	VISIO_GLOW_TRANSPARENCY,
	VISIO_SOFT_EDGE_SIZES,
	VISIO_REFLECTION_PRESETS,
	type VisioGlowEffect,
	type VisioReflectionEffect,
} from './edit-formatting-glow';
export {
	VISIO_BUILT_IN_THEMES,
	VISIO_BUILT_IN_THEME_IDS,
	isVisioBuiltInThemeId,
	visioBuiltInTheme,
	visioThemeVariantColors,
	type VisioBuiltInTheme,
	type VisioBuiltInThemeId,
} from './theme-builtins';
export { visioThemeXml } from './theme-write';
export {
	VISIO_BASIC_SHAPES,
	isVisioBasicShape,
	visioBasicShapeOutline,
	type VisioBasicShape,
	type VisioBasicOutline,
	type VisioOutlinePoint,
} from './basic-shapes';
export { isVisioChangeShapeTarget } from './edit-change-shape-commands';
export { VISIO_SCALE_UNITS } from './edit-page-setup-commands';
export {
	VISIO_BACKGROUND_STYLES,
	VISIO_BORDER_STYLES,
	VISIO_MANAGED_BACKGROUND,
	visioBackgroundShapes,
	visioBorderShapes,
	visioDecorationName,
	visioShade,
	type VisioBackgroundStyle,
	type VisioBorderStyle,
	type VisioDecorationShape,
} from './page-decoration';
export {
	parseVisioFormula,
	analyzeVisioFormula,
	evaluateVisioFormula,
	visioFormulaCachedValue,
	VisioFormulaError,
	type VisioFormulaAst,
	type VisioFormulaLimits,
	type VisioFormulaValue,
	type VisioFormulaReference,
	type VisioFormulaUnit,
} from './formula';
export { visioOpenArrowExtent, visioOpenArrowPath } from './open-arrow';
export {
	visioFilledArrow,
	trimVisioArrowLine,
	layoutVisioFilledArrowLine,
	type VisioFilledArrow,
	type VisioArrowLineLayout,
} from './filled-arrow';
export { VISIO_PICTURE_MAX_BYTES, VISIO_METADATA_TEXT_LIMIT } from './edit-metadata-commands';
