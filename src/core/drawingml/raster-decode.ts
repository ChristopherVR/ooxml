import { unzlibSync } from 'fflate';
import { inspectRasterImage, type RasterImageOptions } from './raster-image';

/** Decoded 8-bit pixels: RGB samples, plus an alpha plane when a pixel is not opaque. */
export interface DecodedRaster {
	width: number;
	height: number;
	rgb: Uint8Array;
	alpha?: Uint8Array;
}

const PNG_CHANNELS: Record<number, number> = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 };

function chunks(bytes: Uint8Array): { type: string; data: Uint8Array }[] {
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const out: { type: string; data: Uint8Array }[] = [];
	for (let at = 8; at + 12 <= bytes.length;) {
		const length = view.getUint32(at);
		const type = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
		if (at + 12 + length > bytes.length) break;
		out.push({ type, data: bytes.subarray(at + 8, at + 8 + length) });
		if (type === 'IEND') break;
		at += 12 + length;
	}
	return out;
}

function paeth(a: number, b: number, c: number): number {
	const p = a + b - c;
	const pa = Math.abs(p - a);
	const pb = Math.abs(p - b);
	const pc = Math.abs(p - c);
	return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
}

/** Reverses the PNG row filters in place, returning the unfiltered rows without filter bytes. */
function unfilter(data: Uint8Array, stride: number, height: number, bpp: number): Uint8Array {
	const out = new Uint8Array(stride * height);
	for (let y = 0; y < height; y++) {
		const filter = data[y * (stride + 1)]!;
		const row = data.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1));
		const at = y * stride;
		for (let x = 0; x < stride; x++) {
			const left = x >= bpp ? out[at + x - bpp]! : 0;
			const up = y ? out[at - stride + x]! : 0;
			const corner = y && x >= bpp ? out[at - stride + x - bpp]! : 0;
			const value = row[x]!;
			out[at + x] =
				(filter === 1
					? value + left
					: filter === 2
						? value + up
						: filter === 3
							? value + ((left + up) >> 1)
							: filter === 4
								? value + paeth(left, up, corner)
								: value) & 0xff;
		}
	}
	return out;
}

/**
 * Decodes a non-interlaced PNG (every colour type, 1 to 8 bits per sample, 16-bit reduced to 8)
 * into RGB and alpha planes. Returns `undefined` for an interlaced or malformed image. The size
 * limits of `inspectRasterImage` apply, and the inflated data is bounded by the declared size.
 */
export function decodePng(
	bytes: Uint8Array,
	limits: RasterImageOptions = {},
): DecodedRaster | undefined {
	let info;
	try {
		info = inspectRasterImage(bytes, limits);
	} catch {
		return undefined;
	}
	if (info.mimeType !== 'image/png') return undefined;
	const parts = chunks(bytes);
	const header = parts.find((part) => part.type === 'IHDR')?.data;
	if (!header || header.length < 13) return undefined;
	const { pixelWidth: width, pixelHeight: height } = info;
	const depth = header[8]!;
	const colorType = header[9]!;
	const channels = PNG_CHANNELS[colorType];
	if (!channels || header[12] !== 0 || ![1, 2, 4, 8, 16].includes(depth)) return undefined;
	const stride = Math.ceil((width * channels * depth) / 8);
	const idat = parts.filter((part) => part.type === 'IDAT');
	const joined = new Uint8Array(idat.reduce((sum, part) => sum + part.data.length, 0));
	let offset = 0;
	for (const part of idat) {
		joined.set(part.data, offset);
		offset += part.data.length;
	}
	let raw: Uint8Array;
	try {
		raw = unzlibSync(joined, { out: new Uint8Array((stride + 1) * height) });
	} catch {
		return undefined;
	}
	if (raw.length < (stride + 1) * height) return undefined;
	const rows = unfilter(raw, stride, height, Math.max(1, Math.ceil((channels * depth) / 8)));
	const sample = (y: number, index: number): number => {
		const at = y * stride;
		if (depth === 8) return rows[at + index]!;
		if (depth === 16) return rows[at + index * 2]!;
		const bit = index * depth;
		const value = (rows[at + (bit >> 3)]! >> (8 - depth - (bit & 7))) & ((1 << depth) - 1);
		return colorType === 3 ? value : Math.round((value * 255) / ((1 << depth) - 1));
	};
	const palette = parts.find((part) => part.type === 'PLTE')?.data;
	const transparency = parts.find((part) => part.type === 'tRNS')?.data;
	const rgb = new Uint8Array(width * height * 3);
	const alpha = new Uint8Array(width * height).fill(255);
	let translucent = false;
	for (let y = 0; y < height; y++)
		for (let x = 0; x < width; x++) {
			const pixel = y * width + x;
			let r: number;
			let g: number;
			let b: number;
			let a = 255;
			if (colorType === 3) {
				const index = sample(y, x);
				if (!palette || index * 3 + 2 >= palette.length) return undefined;
				[r, g, b] = [palette[index * 3]!, palette[index * 3 + 1]!, palette[index * 3 + 2]!];
				if (transparency && index < transparency.length) a = transparency[index]!;
			} else if (colorType === 0 || colorType === 4) {
				r = g = b = sample(y, x * channels);
				if (colorType === 4) a = sample(y, x * channels + 1);
			} else {
				r = sample(y, x * channels);
				g = sample(y, x * channels + 1);
				b = sample(y, x * channels + 2);
				if (colorType === 6) a = sample(y, x * channels + 3);
			}
			rgb[pixel * 3] = r;
			rgb[pixel * 3 + 1] = g;
			rgb[pixel * 3 + 2] = b;
			alpha[pixel] = a;
			if (a !== 255) translucent = true;
		}
	return translucent ? { width, height, rgb, alpha } : { width, height, rgb };
}

/** The number of colour components of a baseline or progressive JPEG (1, 3 or 4), from its frame header. */
export function jpegComponents(bytes: Uint8Array): number | undefined {
	for (let at = 2; at + 9 < bytes.length;) {
		if (bytes[at] !== 0xff) return undefined;
		const marker = bytes[at + 1]!;
		if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
			at += 2;
			continue;
		}
		const length = (bytes[at + 2]! << 8) | bytes[at + 3]!;
		const frame = marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker);
		if (frame) return bytes[at + 9];
		at += 2 + length;
	}
	return undefined;
}
