import { describe, expect, it } from 'vitest';

import { PdfEmbeddedFont } from './pdf-embedded-font';

const be16 = (v: number) => [(v >> 8) & 255, v & 255];
const be32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];

/** A one-glyph TrueType font (only `.notdef`, mapped from nothing) with the given `fsType`. */
function font(fsType: number): Uint8Array {
	const head = [
		...be32(0x00010000),
		...Array(14).fill(0),
		...be16(1000),
		...Array(32).fill(0),
		...be16(0),
		...be16(0),
	];
	const hhea = [...be32(0x00010000), ...Array(30).fill(0), ...be16(1)];
	const maxp = [...be32(0x00010000), ...be16(1), ...Array(26).fill(0)];
	const hmtx = [...be16(500), ...be16(0)];
	const loca = [...be16(0), ...be16(0)];
	const cmapSub = [...be16(4), ...be16(24), ...be16(0), ...be16(2), ...Array(6).fill(0)];
	cmapSub.push(...be16(0xffff), ...be16(0), ...be16(0xffff), ...be16(1), ...be16(0));
	const cmap = [...be16(0), ...be16(1), ...be16(3), ...be16(1), ...be32(12), ...cmapSub];
	const os2 = [...Array(8).fill(0), ...be16(fsType), ...Array(86).fill(0)];
	const tables: Array<[string, number[]]> = [
		['OS/2', os2],
		['cmap', cmap],
		['glyf', []],
		['head', head],
		['hhea', hhea],
		['hmtx', hmtx],
		['loca', loca],
		['maxp', maxp],
	];
	const directory: number[] = [];
	const body: number[] = [];
	let offset = 12 + tables.length * 16;
	for (const [name, bytes] of tables) {
		const padded = [...bytes, ...Array((4 - (bytes.length % 4)) % 4).fill(0)];
		directory.push(...[...name].map((c) => c.charCodeAt(0)), ...be32(0), ...be32(offset));
		directory.push(...be32(bytes.length));
		body.push(...padded);
		offset += padded.length;
	}
	return Uint8Array.from([
		...be32(0x00010000),
		...be16(tables.length),
		...Array(6).fill(0),
		...directory,
		...body,
	]);
}

describe('PdfEmbeddedFont embedding permissions', () => {
	const create = (fsType: number) => PdfEmbeddedFont.create({ family: 'Test', data: font(fsType) });

	it('embeds installable, editable and preview & print faces', () => {
		expect(create(0)).toBeDefined();
		expect(create(0x0008)).toBeDefined();
		expect(create(0x0004)).toBeDefined();
	});

	it('refuses restricted, no-subsetting and bitmap-only faces', () => {
		expect(create(0x0002)).toBeUndefined();
		expect(create(0x0100)).toBeUndefined();
		expect(create(0x0200)).toBeUndefined();
	});
});
