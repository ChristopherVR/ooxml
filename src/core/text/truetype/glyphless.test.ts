import { describe, expect, it } from 'vitest';

import { glyphlessTrueType } from './glyphless';

/** The table directory of an sfnt file (no `cmap` here, so `parseTrueType` does not apply). */
function tables(data: Uint8Array): Map<string, DataView> {
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const out = new Map<string, DataView>();
	for (let index = 0; index < view.getUint16(4); index++) {
		const entry = 12 + index * 16;
		const name = String.fromCharCode(...data.subarray(entry, entry + 4));
		const offset = view.getUint32(entry + 8);
		out.set(name, new DataView(data.buffer, data.byteOffset + offset, view.getUint32(entry + 12)));
	}
	return out;
}

describe('glyphlessTrueType', () => {
	it('writes a TrueType font of empty glyphs with the given advances', () => {
		const data = glyphlessTrueType([0, 1000, 500, 1000]);
		expect(new DataView(data.buffer).getUint32(0)).toBe(0x00010000);
		const font = tables(data);
		expect([...font.keys()]).toEqual(['glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp']);
		expect(font.get('head')!.getUint16(18)).toBe(1000);
		expect(font.get('maxp')!.getUint16(4)).toBe(4);
		expect(font.get('hhea')!.getUint16(34)).toBe(4);
		const hmtx = font.get('hmtx')!;
		expect([0, 1, 2, 3].map((glyph) => hmtx.getUint16(glyph * 4))).toEqual([0, 1000, 500, 1000]);
		// Short loca offsets, all zero: every glyph is empty.
		const loca = font.get('loca')!;
		expect(loca.byteLength).toBe(10);
		for (let i = 0; i < 5; i++) expect(loca.getUint16(i * 2)).toBe(0);
	});

	it('keeps at least the .notdef glyph', () => {
		expect(tables(glyphlessTrueType([])).get('maxp')!.getUint16(4)).toBe(1);
	});

	it('sets the head checksum adjustment so the whole file sums to the magic value', () => {
		const data = glyphlessTrueType([0, 1000]);
		const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
		let sum = 0;
		for (let i = 0; i < data.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
		expect(sum).toBe(0xb1b0afba);
	});
});
