import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { decodeJbig2, decodeJbig2Pages, Jbig2Error } from './jbig2';
import type { Bitmap } from './jbig2-bitmap';

interface Case {
	name: string;
	source: string;
	embedded: boolean;
	globals?: string;
	data: string;
	pages: { width: number; height: number; bits?: string; sha256?: string }[];
}

const cases = JSON.parse(
	readFileSync(new URL('./__fixtures__/jbig2/jbig2.json', import.meta.url), 'utf8'),
) as Case[];
const bytes = (base64: string) => new Uint8Array(Buffer.from(base64, 'base64'));

/** Packed rows, 1 for black (as a PBM holds them). */
function pack(bitmap: Bitmap): Uint8Array {
	const stride = (bitmap.width + 7) >> 3;
	const out = new Uint8Array(stride * bitmap.height);
	for (let y = 0; y < bitmap.height; y++)
		for (let x = 0; x < bitmap.width; x++)
			if (bitmap.data[y * bitmap.width + x]) out[y * stride + (x >> 3)]! |= 0x80 >> (x & 7);
	return out;
}

describe('decodeJbig2Pages (checked against the jbig2dec reference decoder)', () => {
	it.each(cases.map((item) => [`${item.name} (${item.source})`, item] as const))(
		'decodes %s',
		(_name, item) => {
			const pages = decodeJbig2Pages(
				bytes(item.data),
				item.globals ? bytes(item.globals) : undefined,
			);
			expect(pages.map((page) => [page.width, page.height])).toEqual(
				item.pages.map((page) => [page.width, page.height]),
			);
			pages.forEach((page, index) => {
				const expected = item.pages[index]!;
				const packed = pack(page);
				if (expected.bits) expect(Buffer.from(packed).toString('base64')).toBe(expected.bits);
				else expect(createHash('sha256').update(packed).digest('hex')).toBe(expected.sha256);
			});
		},
	);
});

describe('decodeJbig2', () => {
	const generic = cases.find((item) => item.name === 'generic')!;

	it('delivers the first page as PDF rows, 1 for white, cut or padded to the image size', () => {
		const page = generic.pages[0]!;
		const out = decodeJbig2(bytes(generic.data), undefined, page.width, page.height);
		const expected = bytes(page.bits!);
		const stride = (page.width + 7) >> 3;
		const pad = stride * 8 - page.width;
		for (let index = 0; index < out.length; index++) {
			const mask = index % stride === stride - 1 ? 0xff << pad : 0xff;
			expect(out[index]! & mask).toBe(~expected[index]! & mask & 0xff);
		}
		const padded = decodeJbig2(bytes(generic.data), undefined, page.width + 8, page.height + 1);
		expect(padded.length).toBe((stride + 1) * (page.height + 1));
		expect(padded.at(-1)).toBe(0xff);
	});

	it('stops at the size limit, the work limit and damaged data', () => {
		const data = bytes(generic.data);
		expect(() => decodeJbig2(data, undefined, 10, 10, { maxPixels: 1000 })).toThrow(Jbig2Error);
		expect(() => decodeJbig2(data, undefined, 10, 10, { maxWork: 1000 })).toThrow(Jbig2Error);
		expect(() => decodeJbig2(data.subarray(0, 50), undefined, 10, 10)).toThrow(Jbig2Error);
		expect(() => decodeJbig2(new Uint8Array(0), undefined, 10, 10)).toThrow(Jbig2Error);
	});
});
