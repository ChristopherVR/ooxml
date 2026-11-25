import { NS } from '../xml/index';
import { VisioPackage, VisioPackageError } from './package';
import { children, type Report } from './sheet';
import { inspectPicture, inspectPictureInput, type VisioPictureInfo } from './picture-inspection';
import { svgPictureBytes } from '../drawingml/static-svg-picture';
import { tiffPictureToPng } from '../drawingml/tiff-picture';
import type { VisioImage } from './model';

export type { VisioPictureInfo } from './picture-inspection';

export type {
	RasterImageOptions as VisioImageOptions,
	RasterImageInfo as VisioRasterImageInfo,
} from '../drawingml/raster-image';
import {
	inspectRasterImage,
	RasterImageError,
	RASTER_IMAGE_LIMITS,
	type RasterImageOptions as VisioImageOptions,
	type RasterImageInfo as VisioRasterImageInfo,
	type RasterImageErrorCode as VisioImageErrorCode,
} from '../drawingml/raster-image';
export type { RasterImageErrorCode as VisioImageErrorCode } from '../drawingml/raster-image';
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

/** Inspect a local picture by bytes, before constructing an insert command. */
export function inspectVisioPicture(bytes: Uint8Array): VisioPictureInfo {
	try {
		const { mimeType, pixelWidth, pixelHeight } = inspectPictureInput(bytes);
		return { mimeType, pixelWidth, pixelHeight };
	} catch (error) {
		if (error instanceof RasterImageError) throw new VisioImageError(error.code, error.message);
		throw error;
	}
}

/** Read only authenticated internal image parts; no URLs or active media are returned. */
export async function readVisioImage(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
	options: VisioImageOptions = {},
): Promise<VisioImage | undefined> {
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
		cache = { assets: new Map(), tiffPixels: 0, tiffPreviewBytes: 0 };
		validationCache.set(pkg, cache);
	}
	const key = `${rel.target}\0${limits.maxImageBytes}/${limits.maxPixels}/${limits.maxDimension}`;
	let pending = cache.assets.get(key);
	if (!pending) {
		pending = prepareImage(bytes, limits, cache, rel.target);
		cache.assets.set(key, pending);
	}
	try {
		return await pending;
	} catch (error) {
		if (error instanceof RasterImageError) return omit(error.code, error.message);
		throw error;
	}
}

interface ImageCache {
	assets: Map<string, Promise<VisioImage>>;
	tiffPixels: number;
	tiffPreviewBytes: number;
}
// Package-keyed cache is weak, so finished imports do not retain package buffers.
const validationCache = new WeakMap<VisioPackage, ImageCache>();

async function prepareImage(
	bytes: Uint8Array,
	limits: ImageLimits,
	cache: ImageCache,
	target: string,
): Promise<VisioImage> {
	const info = inspectPicture(bytes, limits, false);
	const { pixelWidth, pixelHeight } = info;
	if (info.mimeType === 'image/svg+xml') {
		if (!target.toLowerCase().endsWith('.svg'))
			throw new RasterImageError('unsupported-image', 'SVG requires an SVG image part.');
		return { mimeType: info.mimeType, pixelWidth, pixelHeight, bytes: svgPictureBytes(info.svg) };
	}
	if (info.mimeType === 'image/tiff') {
		cache.tiffPixels += pixelWidth * pixelHeight;
		if (cache.tiffPixels > 32_000_000)
			throw new RasterImageError(
				'image-limit',
				'Document TIFF previews exceed the decoded pixel limit.',
			);
		const preview = await tiffPictureToPng(info.tiff);
		cache.tiffPreviewBytes += preview.length;
		if (
			preview.length > VISIO_RASTER_IMAGE_LIMITS.maxImageBytes ||
			cache.tiffPreviewBytes > 32 * 1024 * 1024
		)
			throw new RasterImageError('image-limit', 'TIFF previews exceed the PNG byte limit.');
		return { mimeType: 'image/png', pixelWidth, pixelHeight, bytes: preview };
	}
	if (info.mimeType === 'image/x-emf')
		throw new RasterImageError('unsupported-image', 'Expected a bitmap.');
	return { mimeType: info.mimeType, pixelWidth, pixelHeight, bytes };
}
