/** Read-only VSDX parsing and normalized scenes. Legacy VSD/VDX and saving are not supported. */
export type * from './model.js';
export type { VisioShapeData, VisioHyperlink, VisioMetadataOptions } from './shape-metadata.js';
export { parseVsdx, getVisioPageLayers, type ParseVsdxOptions } from './parser.js';
export { VisioPackageError, type VisioPackageLimits } from './package.js';
export {
	resolveVisioPageVisibility,
	VISIO_VISIBILITY_LIMITS,
	type VisioVisibilityOptions,
	type VisioResolvedShapeVisibility,
} from './visibility.js';
export {
	inspectVisioRasterImage,
	VisioImageError,
	VISIO_RASTER_IMAGE_LIMITS,
	type VisioImageOptions,
	type VisioImageErrorCode,
	type VisioRasterImageInfo,
} from './media.js';
export {
	inspectVisioEmfAdmission,
	VISIO_EMF_ADMISSION_LIMITS,
	type VisioEmfAdmissionOptions,
	type VisioEmfAdmissionResult,
	type VisioEmfAdmissionStatus,
	type VisioEmfAdmissionDiagnostic,
	type VisioEmfAdmissionMetrics,
} from './emf-admission.js';
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
} from './foreign-vector.js';
export {
	convertVisioMetafile,
	type VisioMetafileTreeConverter,
	type VisioMetafileConversionResult,
} from './convert-metafile.js';
