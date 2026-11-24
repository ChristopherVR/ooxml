import { zlibSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { crc32 } from '../digest/crc32';
import { decodePng, jpegComponents } from './raster-decode';

function chunk(type: string, data: Uint8Array): Uint8Array {
	const body = new Uint8Array(4 + data.length);
	body.set([...type].map((c) => c.charCodeAt(0)));
	body.set(data, 4);
	const out = new Uint8Array(12 + data.length);
	const view = new DataView(out.buffer);
	view.setUint32(0, data.length);
	out.set(body, 4);
	view.setUint32(8 + data.length, (crc32(body) ^ 0xffffffff) >>> 0);
	return out;
}

/** A PNG of `rows` (each already prefixed with its filter byte). */
function png(
	width: number,
	height: number,
	depth: number,
	colorType: number,
	rows: number[][],
	extra: Uint8Array[] = [],
): Uint8Array {
	const header = new Uint8Array(13);
	const view = new DataView(header.buffer);
	view.setUint32(0, width);
	view.setUint32(4, height);
	header.set([depth, colorType, 0, 0, 0], 8);
	const parts = [
		new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
		chunk('IHDR', header),
		...extra,
		chunk('IDAT', zlibSync(new Uint8Array(rows.flat()))),
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

describe('decodePng', () => {
	it('decodes RGBA with Sub and Up filters and keeps the alpha plane', () => {
		const bytes = png(2, 2, 8, 6, [
			[1, 10, 20, 30, 200, 5, 5, 5, 55],
			[2, 1, 1, 1, 0, 0, 0, 0, 0],
		]);
		const decoded = decodePng(bytes)!;
		expect([...decoded.rgb]).toEqual([10, 20, 30, 15, 25, 35, 11, 21, 31, 15, 25, 35]);
		expect([...decoded.alpha!]).toEqual([200, 255, 200, 255]);
	});

	it('drops the alpha plane of an opaque image and expands 1-bit grey', () => {
		const decoded = decodePng(png(3, 1, 1, 0, [[0, 0b10100000]]))!;
		expect([...decoded.rgb]).toEqual([255, 255, 255, 0, 0, 0, 255, 255, 255]);
		expect(decoded.alpha).toBeUndefined();
	});

	it('maps palette indices and their transparency', () => {
		const decoded = decodePng(
			png(
				2,
				1,
				8,
				3,
				[[0, 1, 0]],
				[
					chunk('PLTE', new Uint8Array([255, 0, 0, 0, 0, 255])),
					chunk('tRNS', new Uint8Array([128])),
				],
			),
		)!;
		expect([...decoded.rgb]).toEqual([0, 0, 255, 255, 0, 0]);
		expect([...decoded.alpha!]).toEqual([255, 128]);
	});

	it('refuses interlaced, truncated and non-PNG input', () => {
		const interlaced = png(1, 1, 8, 2, [[0, 1, 2, 3]]);
		interlaced[8 + 8 + 12] = 1;
		expect(decodePng(interlaced)).toBeUndefined();
		expect(decodePng(new Uint8Array([1, 2, 3]))).toBeUndefined();
	});
});

describe('jpegComponents', () => {
	it('reads the component count from the frame header', () => {
		const jpeg = new Uint8Array([
			0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 11, 8, 0, 1, 0, 1, 3, 1, 0x11, 0, 0, 0,
		]);
		expect(jpegComponents(jpeg)).toBe(3);
		expect(jpegComponents(new Uint8Array([0xff, 0xd8, 0x00]))).toBeUndefined();
	});
});
