import { zlibSync } from 'fflate';
import { crc32 } from '../digest/crc32';

function chunk(type: string, data: Uint8Array): Uint8Array {
	const out = new Uint8Array(12 + data.length);
	const view = new DataView(out.buffer);
	view.setUint32(0, data.length);
	for (let index = 0; index < 4; index++) out[4 + index] = type.charCodeAt(index);
	out.set(data, 8);
	view.setUint32(8 + data.length, (crc32(out.subarray(4, 8 + data.length)) ^ 0xffffffff) >>> 0);
	return out;
}

/**
 * An 8-bit PNG of `rgb` samples (`width * height * 3` bytes), RGBA when an `alpha` plane is given
 * (`width * height` bytes). Rows use the Up filter, which suits scanned and drawn pictures alike.
 */
export function encodePng(
	width: number,
	height: number,
	rgb: Uint8Array,
	alpha?: Uint8Array,
): Uint8Array {
	const channels = alpha ? 4 : 3;
	const stride = width * channels;
	const raw = new Uint8Array((stride + 1) * height);
	let previous = new Uint8Array(stride);
	const row = new Uint8Array(stride);
	for (let y = 0; y < height; y++) {
		for (let x = 0; x < width; x++) {
			const pixel = y * width + x;
			row[x * channels] = rgb[pixel * 3] ?? 0;
			row[x * channels + 1] = rgb[pixel * 3 + 1] ?? 0;
			row[x * channels + 2] = rgb[pixel * 3 + 2] ?? 0;
			if (alpha) row[x * channels + 3] = alpha[pixel] ?? 255;
		}
		const at = y * (stride + 1);
		raw[at] = 2;
		for (let index = 0; index < stride; index++)
			raw[at + 1 + index] = (row[index]! - previous[index]!) & 0xff;
		previous = row.slice();
	}
	const header = new Uint8Array(13);
	const view = new DataView(header.buffer);
	view.setUint32(0, width);
	view.setUint32(4, height);
	header.set([8, alpha ? 6 : 2, 0, 0, 0], 8);
	const parts = [
		Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		chunk('IDAT', zlibSync(raw)),
		chunk('IEND', new Uint8Array()),
	];
	const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
	let at = 0;
	for (const part of parts) {
		out.set(part, at);
		at += part.length;
	}
	return out;
}
