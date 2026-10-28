import { crc32 } from '../digest/crc32';

export interface RasterImageOptions {
	maxImageBytes?: number;
	maxPixels?: number;
	maxDimension?: number;
}
export interface RasterImageInfo {
	mimeType: 'image/png' | 'image/jpeg' | 'image/gif';
	pixelWidth: number;
	pixelHeight: number;
}
type ImageLimits = Required<RasterImageOptions>;
/** Defaults and hard ceilings for inspecting caller-owned raster bytes. */
export const RASTER_IMAGE_LIMITS: Readonly<ImageLimits> = Object.freeze({
	maxImageBytes: 8 * 1024 * 1024,
	maxPixels: 16_000_000,
	maxDimension: 16384,
});
export type RasterImageErrorCode =
	| 'invalid-image'
	| 'unsupported-image'
	| 'image-limit'
	| 'invalid-limit';
export class RasterImageError extends Error {
	constructor(
		public readonly code: RasterImageErrorCode,
		message: string,
	) {
		super(message);
		this.name = 'RasterImageError';
	}
}
const invalid = (message: string): never => {
	throw new RasterImageError('invalid-image', message);
};
const unsupported = (message: string): never => {
	throw new RasterImageError('unsupported-image', message);
};
function dimensions(width: number, height: number, limits: ImageLimits): void {
	if (!width || !height) invalid('Raster image has zero dimensions.');
	if (
		width > limits.maxDimension ||
		height > limits.maxDimension ||
		width * height > limits.maxPixels
	)
		throw new RasterImageError(
			'image-limit',
			'Raster image dimensions exceed the configured limit.',
		);
}
const ascii = (bytes: Uint8Array, start: number, length: number) =>
	String.fromCharCode(...bytes.subarray(start, start + length));
function png(bytes: Uint8Array, limits: ImageLimits): [number, number] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	if (bytes.length < 33 || view.getUint32(8) !== 13 || ascii(bytes, 12, 4) !== 'IHDR')
		return invalid('PNG is missing a complete IHDR header.');
	const width = view.getUint32(16),
		height = view.getUint32(20);
	dimensions(width, height, limits);
	const depth = bytes[24]!,
		color = bytes[25]!;
	const depths: Record<number, number[]> = {
		0: [1, 2, 4, 8, 16],
		2: [8, 16],
		3: [1, 2, 4, 8],
		4: [8, 16],
		6: [8, 16],
	};
	if (!depths[color]?.includes(depth) || bytes[26] !== 0 || bytes[27] !== 0 || bytes[28]! > 1)
		invalid('PNG header uses invalid color, compression, filter, or interlace values.');
	let data = false;
	for (let at = 8; at < bytes.length;) {
		if (at + 12 > bytes.length) invalid('Truncated PNG chunk.');
		const length = view.getUint32(at),
			type = ascii(bytes, at + 4, 4),
			end = at + length + 12;
		if (end > bytes.length) invalid('Truncated PNG chunk payload.');
		if ((crc32(bytes.subarray(at + 4, end - 4)) ^ 0xffffffff) >>> 0 !== view.getUint32(end - 4))
			invalid('PNG chunk checksum mismatch.');
		if (type === 'acTL') unsupported('Animated PNG is not supported.');
		if (type === 'IHDR' && at !== 8) invalid('PNG contains multiple image headers.');
		if (type === 'IDAT' && length) data = true;
		if (type === 'IEND') {
			if (length || end !== bytes.length || !data) invalid('PNG has an invalid image terminator.');
			return [width, height];
		}
		at = end;
	}
	return invalid('PNG has no image terminator.');
}
function gif(bytes: Uint8Array, limits: ImageLimits): [number, number] {
	if (bytes.length < 13) return invalid('GIF is missing a complete screen descriptor.');
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const width = view.getUint16(6, true),
		height = view.getUint16(8, true);
	dimensions(width, height, limits);
	let at = 13 + (bytes[10]! & 128 ? 3 * 2 ** ((bytes[10]! & 7) + 1) : 0);
	let frames = 0,
		pixels = 0;
	const blocks = () => {
		for (;;) {
			if (at >= bytes.length) invalid('GIF has truncated data blocks.');
			const size = bytes[at++]!;
			if (!size) return;
			at += size;
			if (at > bytes.length) invalid('GIF has a truncated data block.');
		}
	};
	while (at < bytes.length) {
		const marker = bytes[at++]!;
		if (marker === 0x3b) {
			if (!frames || at !== bytes.length) invalid('GIF has an invalid image terminator.');
			return [width, height];
		}
		if (marker === 0x21) {
			if (at >= bytes.length) invalid('GIF has a truncated extension.');
			at++;
			blocks();
			continue;
		}
		if (marker !== 0x2c || at + 9 > bytes.length) invalid('GIF has an invalid image descriptor.');
		const left = view.getUint16(at, true),
			top = view.getUint16(at + 2, true);
		const frameWidth = view.getUint16(at + 4, true),
			frameHeight = view.getUint16(at + 6, true);
		dimensions(frameWidth, frameHeight, limits);
		if (left + frameWidth > width || top + frameHeight > height)
			invalid('GIF frame exceeds its screen.');
		// Bound animation work as well as the logical canvas size.
		pixels += width * height;
		if (pixels > limits.maxPixels)
			throw new RasterImageError('image-limit', 'GIF animation exceeds the pixel limit.');
		const packed = bytes[at + 8]!;
		at += 9 + (packed & 128 ? 3 * 2 ** ((packed & 7) + 1) : 0);
		if (at >= bytes.length || bytes[at]! < 2 || bytes[at]! > 8)
			invalid('GIF has an invalid LZW header.');
		at++;
		blocks();
		frames++;
	}
	return invalid('GIF has no image terminator.');
}
function jpeg(bytes: Uint8Array, limits: ImageLimits): [number, number] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	let at = 2,
		width = 0,
		height = 0,
		scan = false;
	while (at < bytes.length) {
		if (bytes[at++] !== 0xff) invalid('JPEG has an invalid segment marker.');
		while (bytes[at] === 0xff) at++;
		const marker = bytes[at++];
		if (marker === 0xd9) {
			if (!width || !scan || at !== bytes.length) invalid('JPEG has an invalid image terminator.');
			return [width, height];
		}
		if (
			marker === undefined ||
			marker === 0 ||
			marker === 0xd8 ||
			marker === 1 ||
			(marker >= 0xd0 && marker <= 0xd7)
		)
			invalid('JPEG has an unexpected standalone marker.');
		if (at + 2 > bytes.length) invalid('JPEG has a truncated segment header.');
		const length = view.getUint16(at),
			end = at + length;
		if (length < 2 || end > bytes.length) invalid('JPEG has a truncated segment.');
		if (marker! >= 0xc0 && marker! <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker!)) {
			if (![0xc0, 0xc1, 0xc2].includes(marker!))
				unsupported('This JPEG frame coding is unsupported.');
			if (width || length < 11 || bytes[at + 2] !== 8 || length !== 8 + 3 * bytes[at + 7]!)
				invalid('JPEG has an invalid frame header.');
			height = view.getUint16(at + 3);
			width = view.getUint16(at + 5);
			dimensions(width, height, limits);
		}
		if (marker === 0xdc) unsupported('JPEG deferred dimensions are unsupported.');
		at = end;
		if (marker === 0xda) {
			if (!width || length < 8 || length !== 6 + 2 * bytes[end - length + 2]!)
				invalid('JPEG has an invalid scan header.');
			scan = true;
			// Entropy bytes are not decoded; skip escaped FF and restart markers safely.
			while (at < bytes.length) {
				if (bytes[at] !== 0xff) {
					at++;
					continue;
				}
				const next = bytes[at + 1];
				if (next === 0 || (next !== undefined && next >= 0xd0 && next <= 0xd7)) {
					at += 2;
					continue;
				}
				break;
			}
		}
	}
	return invalid('JPEG has no image terminator.');
}

/**
 * Inspect PNG, JPEG, or GIF structure and intrinsic dimensions without trusting declared metadata.
 * Checks PNG chunk checksums, but does not decode pixel/entropy data or prove browser decodability.
 * Limits are positive safe integers and may only lower RASTER_IMAGE_LIMITS. Throws
 * RasterImageError on invalid limits, unsupported/malformed input, or an exceeded limit.
 * Reads only the supplied view and does not cache, copy, or retain mutable caller-owned bytes.
 */
export function inspectRasterImage(
	bytes: Uint8Array,
	options: RasterImageOptions = {},
): RasterImageInfo {
	if (!options || typeof options !== 'object' || Array.isArray(options))
		throw new RasterImageError('invalid-limit', 'Image limits must be an options object.');
	const limits = { ...RASTER_IMAGE_LIMITS, ...options };
	for (const key of Object.keys(RASTER_IMAGE_LIMITS) as (keyof ImageLimits)[]) {
		const value = limits[key];
		if (!Number.isSafeInteger(value) || value <= 0 || value > RASTER_IMAGE_LIMITS[key])
			throw new RasterImageError('invalid-limit', `Invalid raster image limit: ${key}.`);
	}
	if (!(bytes instanceof Uint8Array)) invalid('Raster image bytes must be a Uint8Array.');
	if (bytes.byteLength > limits.maxImageBytes)
		throw new RasterImageError('image-limit', 'Raster image bytes exceed the configured limit.');
	return inspectRasterImageStructure(bytes, limits);
}

export function inspectRasterImageStructure(
	bytes: Uint8Array,
	limits: ImageLimits,
): RasterImageInfo {
	let mimeType: RasterImageInfo['mimeType'], size: [number, number];
	if ([137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)) {
		mimeType = 'image/png';
		size = png(bytes, limits);
	} else if (['GIF87a', 'GIF89a'].includes(ascii(bytes, 0, 6))) {
		mimeType = 'image/gif';
		size = gif(bytes, limits);
	} else if (bytes[0] === 0xff && bytes[1] === 0xd8) {
		mimeType = 'image/jpeg';
		size = jpeg(bytes, limits);
	} else return unsupported('Foreign data is not a supported PNG, JPEG, or GIF raster image.');
	return { mimeType, pixelWidth: size[0], pixelHeight: size[1] };
}
