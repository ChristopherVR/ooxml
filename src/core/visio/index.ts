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
	type VisioPageInsert,
	type VisioPageReorder,
	type VisioPageRename,
	type VisioPageDelete,
	type VisioPageEdit,
	type VisioFormatEdit,
	type VisioTextFormatEdit,
	type VisioShapeFormatEdit,
	type VisioShapeOrderEdit,
	type VisioDuplicateShapesEdit,
	type VisioPasteShapesEdit,
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
