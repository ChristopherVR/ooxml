import { checkImageDimensions, RasterImageError, type ImageLimits } from './raster-image';
import { encodePng } from './raster-encode';
import { tiffLzw, tiffPackBits } from './tiff-compression';

const TAG_SIZES = [0, 1, 1, 2, 4, 8, 1, 1, 2, 4, 8, 4, 8, 4] as const;
const TIFF_TAGS = new Set([
	256, 257, 258, 259, 262, 266, 273, 274, 277, 278, 279, 284, 317, 320, 338, 339,
]);
const MAX_STRIPS = 2048;
const invalid = (message: string): never => {
	throw new RasterImageError('invalid-image', message);
};
const unsupported = (): never => {
	throw new RasterImageError('unsupported-image', 'Unsupported TIFF layout or pixel format.');
};
export function hasTiffSignature(view: DataView): boolean {
	return view.byteLength >= 4 && [0x49492a00, 0x4d4d002a].includes(view.getUint32(0));
}

function readTags(view: DataView): Map<number, number[]> {
	if (view.byteLength < 8) invalid('Truncated TIFF header.');
	const little = view.getUint16(0) === 0x4949;
	const word = (at: number) => view.getUint16(at, little);
	const long = (at: number) => view.getUint32(at, little);
	const offset = long(4);
	if (offset < 8 || offset > view.byteLength - 2) invalid('Invalid TIFF directory offset.');
	const count = word(offset);
	const end = offset + 2 + count * 12;
	if (end + 4 > view.byteLength) invalid('Truncated TIFF directory.');
	if (count > 256 || long(end) !== 0) unsupported();
	const tags = new Map<number, number[]>();
	for (let at = offset + 2; at < end; at += 12) {
		const tag = word(at);
		const type = word(at + 2);
		const size = TAG_SIZES[type];
		const items = long(at + 4);
		const length = (size ?? 0) * items;
		const start = length > 4 ? long(at + 8) : at + 8;
		if (!size || !items || start < 8 || length > view.byteLength - start)
			invalid('Invalid TIFF directory field range.');
		if (!TIFF_TAGS.has(tag)) continue;
		if ((type !== 3 && type !== 4) || items > MAX_STRIPS || tags.has(tag))
			invalid('Invalid TIFF image field.');
		tags.set(
			tag,
			Array.from({ length: items }, (_, i) =>
				type === 3 ? word(start + i * 2) : long(start + i * 4),
			),
		);
	}
	return tags;
}

export interface TiffPicture {
	pixelWidth: number;
	pixelHeight: number;
	samples: number;
	bits: number[];
	photometric: number;
	palette: number[];
	predictor: number;
	alpha: number;
	stride: number;
	strips: { encoded: Uint8Array; decodedSize: number }[];
	compression: number;
}

/** Validate a single-strip-layout TIFF image and every compressed strip before decoding pixels. */
export function inspectTiff(view: DataView, limits: ImageLimits): TiffPicture {
	const tags = readTags(view);
	const single = (tag: number, fallback = 0): number => {
		const entry = tags.get(tag);
		if (!entry) return fallback;
		if (entry.length !== 1) invalid('Expected a scalar TIFF image field.');
		return entry[0]!;
	};
	const pixelWidth = single(256);
	const pixelHeight = single(257);
	checkImageDimensions(pixelWidth, pixelHeight, limits);
	const samples = single(277, 1);
	const bits = tags.get(258) ?? [1];
	const compression = single(259, 1);
	const photometric = single(262, -1);
	const predictor = single(317, 1);
	const alpha = single(338, 0);
	const palette = tags.get(320) ?? [];
	const depth = bits[0]!;
	const sampleFormat = tags.get(339) ?? [1];
	const supportedColor =
		(photometric === 0 && samples === 1 && [1, 4, 8].includes(depth)) ||
		(photometric === 1 && samples === 1 && [1, 2, 8].includes(depth)) ||
		(photometric === 2 &&
			depth === 8 &&
			(samples === 3 || (samples === 4 && [1, 2].includes(alpha)))) ||
		(photometric === 3 && samples === 1 && depth === 8 && palette.length === 768);
	if (
		!supportedColor ||
		bits.length !== samples ||
		bits.some((value) => value !== depth) ||
		sampleFormat.some((value) => value !== 1) ||
		single(284, 1) !== 1 ||
		single(266, 1) !== 1 ||
		single(274, 1) !== 1 ||
		![1, 5, 32773].includes(compression) ||
		(predictor !== 1 && (predictor !== 2 || depth !== 8))
	)
		unsupported();
	const offsets = tags.get(273) ?? [];
	const sizes = tags.get(279) ?? [];
	const rows = single(278, pixelHeight);
	const stride = Math.ceil((pixelWidth * samples * depth) / 8);
	if (!rows || offsets.length !== Math.ceil(pixelHeight / rows) || offsets.length !== sizes.length)
		invalid('Invalid TIFF strip count.');
	const ranges = offsets.map((start, i) => ({ start, end: start + sizes[i]! }));
	ranges.sort((a, b) => a.start - b.start);
	let end = 8;
	for (const range of ranges) {
		if (range.start < end || range.end === range.start || range.end > view.byteLength)
			invalid('Invalid or overlapping TIFF strip range.');
		end = range.end;
	}
	const strips = offsets.map((offset, i) => {
		const encoded = new Uint8Array(view.buffer, view.byteOffset + offset, sizes[i]);
		const decodedSize = stride * Math.min(rows, pixelHeight - i * rows);
		if (compression === 5) tiffLzw(encoded, decodedSize);
		else if (compression === 32773) tiffPackBits(encoded, decodedSize);
		else if (encoded.length !== decodedSize) invalid('Truncated or oversized TIFF pixel strip.');
		return { encoded, decodedSize };
	});
	return {
		pixelWidth,
		pixelHeight,
		samples,
		bits,
		photometric,
		palette,
		predictor,
		alpha,
		stride,
		strips,
		compression,
	};
}

/** Reuse PowerPoint's UTIF pixel conversion after decoding only validated strip data. */
export async function tiffPictureToPng(picture: TiffPicture): Promise<Uint8Array> {
	const { toRGBA8 } = await import('utif').catch(() => {
		throw new RasterImageError('unsupported-image', 'The TIFF preview decoder is unavailable.');
	});
	const { pixelWidth: width, pixelHeight: height, stride, samples } = picture;
	const data = new Uint8Array(stride * height);
	let at = 0;
	for (const { encoded, decodedSize } of picture.strips) {
		const output = data.subarray(at, at + decodedSize);
		if (picture.compression === 5) tiffLzw(encoded, decodedSize, output);
		else if (picture.compression === 32773) tiffPackBits(encoded, decodedSize, output);
		else output.set(encoded);
		at += decodedSize;
	}
	if (picture.predictor === 2)
		for (let row = 0; row < data.length; row += stride)
			for (let i = row + samples; i < row + stride; i++) data[i] = data[i]! + data[i - samples]!;
	const rgba = toRGBA8({
		width,
		height,
		data,
		t258: picture.bits,
		t262: [picture.photometric],
		t320: picture.palette,
	});
	const rgb = new Uint8Array(width * height * 3);
	const alpha = samples === 4 ? new Uint8Array(width * height) : undefined;
	for (let i = 0; i < width * height; i++) {
		const a = rgba[i * 4 + 3]!;
		const factor = picture.alpha === 1 && a > 0 ? 255 / a : 1;
		for (let channel = 0; channel < 3; channel++)
			rgb[i * 3 + channel] = Math.min(255, Math.round(rgba[i * 4 + channel]! * factor));
		if (alpha) alpha[i] = a;
	}
	return encodePng(width, height, rgb, alpha);
}
