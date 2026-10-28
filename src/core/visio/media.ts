import { NS } from '../xml/index';
import { VisioPackage, VisioPackageError } from './package';
import { children, type Report } from './sheet';

export type {
	RasterImageOptions as VisioImageOptions,
	RasterImageInfo as VisioRasterImageInfo,
} from '../drawingml/raster-image';
import {
	inspectRasterImage,
	inspectRasterImageStructure,
	RasterImageError,
	RASTER_IMAGE_LIMITS,
	type RasterImageOptions as VisioImageOptions,
	type RasterImageInfo as VisioRasterImageInfo,
	type RasterImageErrorCode as VisioImageErrorCode,
} from '../drawingml/raster-image';
export type { RasterImageErrorCode as VisioImageErrorCode } from '../drawingml/raster-image';
export interface VisioRasterImage extends VisioRasterImageInfo {
	bytes: Uint8Array;
}
type ImageLimits = Required<VisioImageOptions>;
export const VISIO_RASTER_IMAGE_LIMITS = RASTER_IMAGE_LIMITS;
export class VisioImageError extends Error {
	constructor(
		public readonly code: VisioImageErrorCode,
		message: string,
	) {
		super(message);
		this.name = 'VisioImageError';
	}
}
function inspect(bytes: Uint8Array, limits: ImageLimits): VisioRasterImageInfo {
	try {
		return inspectRasterImageStructure(bytes, limits);
	} catch (error) {
		if (error instanceof RasterImageError) throw new VisioImageError(error.code, error.message);
		throw error;
	}
}
export function inspectVisioRasterImage(
	bytes: Uint8Array,
	options: VisioImageOptions = {},
): VisioRasterImageInfo {
	try {
		return inspectRasterImage(bytes, options);
	} catch (error) {
		if (error instanceof RasterImageError) throw new VisioImageError(error.code, error.message);
		throw error;
	}
}

/** Read only authenticated internal raster parts; no URLs or active media are returned. */
export async function readVisioImage(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
	options: VisioImageOptions = {},
): Promise<VisioRasterImage | undefined> {
	const limits: ImageLimits = {
		...VISIO_RASTER_IMAGE_LIMITS,
		...options,
	};
	if (Object.values(limits).some((value) => !Number.isSafeInteger(value) || value <= 0))
		throw new VisioPackageError('INVALID_LIMIT', 'Image limits must be positive safe integers.');
	const omit = (code: string, message: string) => {
		report(code, message, { part: sourcePart });
		return undefined;
	};
	if (
		foreignData.getAttribute('ForeignType') &&
		foreignData.getAttribute('ForeignType') !== 'Bitmap'
	)
		return omit('unsupported-image', 'Only embedded bitmap foreign data is supported.');
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel)
		return omit('invalid-image-relationship', 'Image has no unambiguous relationship reference.');
	if (rel.mode !== 'Internal')
		return omit('external-image', 'External image relationships are not loaded.');
	if (rel.type !== `${NS.r}/image`)
		return omit(
			'invalid-image-relationship',
			'Foreign data does not reference an image relationship.',
		);
	if (pkg.getPartByteLength(rel.target) > limits.maxImageBytes)
		return omit('image-limit', 'Raster image bytes exceed the configured limit before inflation.');
	const bytes = await pkg.readBytes(rel.target);
	if (bytes.length > limits.maxImageBytes)
		return omit('image-limit', 'Raster image bytes exceed the configured limit.');
	let cache = validationCache.get(pkg);
	if (!cache) {
		cache = new Map();
		validationCache.set(pkg, cache);
	}
	const key = `${rel.target}\0${limits.maxImageBytes}/${limits.maxPixels}/${limits.maxDimension}`;
	let result = cache.get(key);
	if (!result) {
		try {
			result = { ...inspect(bytes, limits), bytes };
		} catch (error) {
			if (!(error instanceof VisioImageError)) throw error;
			result = error;
		}
		cache.set(key, result);
	}
	return result instanceof VisioImageError ? omit(result.code, result.message) : result;
}

// Package-keyed cache is weak, so finished imports do not retain package buffers.
const validationCache = new WeakMap<
	VisioPackage,
	Map<string, VisioRasterImage | VisioImageError>
>();
