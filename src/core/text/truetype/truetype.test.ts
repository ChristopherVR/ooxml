import { existsSync, readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseTrueType } from './parse';
import { subsetTrueType } from './subset';

function table(name: string, bytes: number[]) {
	return { name, bytes };
}
const be16 = (v: number) => [(v >> 8) & 255, v & 255];
const be32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];

/** A four-glyph font: .notdef, A (simple), B (composite of A) and C (blank, a space). */
function syntheticFont(): Uint8Array {
	const simple = [...be16(1), ...be16(0), ...be16(0), ...be16(100), ...be16(100), 0, 0, 0, 0, 0, 0];
	const composite = [
		...be16(0xffff),
		...be16(0),
		...be16(0),
		...be16(100),
		...be16(100),
		...be16(0x0003),
		...be16(1),
		...be16(0),
		...be16(0),
	];
	const pad = (b: number[]) => (b.length % 4 ? [...b, ...Array(4 - (b.length % 4)).fill(0)] : b);
	const glyphs = [[], pad(simple), pad(composite), []] as number[][];
	const glyf = glyphs.flat();
	let at = 0;
	const loca = [0, ...glyphs.map((g) => (at += g.length))].flatMap(be32);
	const head = [
		...be32(0x00010000),
		...be32(0),
		...be32(0),
		...be32(0x5f0f3cf5),
		...be16(0),
		...be16(1000),
		...Array(16).fill(0),
		...be16(0),
		...be16(0),
		...be16(100),
		...be16(100),
		...be16(0),
		...be16(8),
		...be16(2),
		...be16(1),
		...be16(0),
	];
	const hhea = [...be32(0x00010000), ...be16(800), ...be16(-200), ...Array(26).fill(0), ...be16(4)];
	const maxp = [...be32(0x00010000), ...be16(4), ...Array(26).fill(0)];
	const hmtx = [500, 600, 700, 300].flatMap((w) => [...be16(w), ...be16(10)]);
	// cmap format 4: A..B -> glyphs 1..2 (delta -64), space -> 3 (delta -29), end 0xFFFF.
	const segments = [
		[0x20, 0x20, 3 - 0x20],
		[0x41, 0x42, 1 - 0x41],
		[0xffff, 0xffff, 1],
	] as const;
	const sub = [
		...be16(4),
		...be16(16 + segments.length * 8),
		...be16(0),
		...be16(segments.length * 2),
		...be16(0),
		...be16(0),
		...be16(0),
		...segments.flatMap((s) => be16(s[1])),
		...be16(0),
		...segments.flatMap((s) => be16(s[0])),
		...segments.flatMap((s) => be16(s[2] & 0xffff)),
		...segments.flatMap(() => be16(0)),
	];
	const cmap = [...be16(0), ...be16(1), ...be16(3), ...be16(1), ...be32(12), ...sub];
	const tables = [
		table('cmap', cmap),
		table('glyf', glyf),
		table('head', head),
		table('hhea', hhea),
		table('hmtx', hmtx),
		table('loca', loca),
		table('maxp', maxp),
	];
	const header = [...be32(0x00010000), ...be16(tables.length), ...Array(6).fill(0)];
	let offset = 12 + tables.length * 16;
	const directory: number[] = [];
	const body: number[] = [];
	for (const t of tables) {
		directory.push(...[...t.name].map((c) => c.charCodeAt(0)), ...be32(0), ...be32(offset));
		directory.push(...be32(t.bytes.length));
		const padded =
			t.bytes.length % 4 ? [...t.bytes, ...Array(4 - (t.bytes.length % 4)).fill(0)] : t.bytes;
		body.push(...padded);
		offset += padded.length;
	}
	return Uint8Array.from([...header, ...directory, ...body]);
}

describe('TrueType parsing', () => {
	it('reads metrics, the character map and outlines', () => {
		const font = parseTrueType(syntheticFont())!;
		expect(font.unitsPerEm).toBe(1000);
		expect(font.glyphCount).toBe(4);
		expect(font.glyphFor(0x41)).toBe(1);
		expect(font.glyphFor(0x42)).toBe(2);
		expect(font.glyphFor(0x20)).toBe(3);
		expect(font.glyphFor(0x43)).toBe(0);
		expect(font.advance(2)).toBe(700);
		expect(font.outline(0).length).toBe(0);
		expect(font.outline(1).length).toBeGreaterThan(0);
	});

	it('refuses data that is not TrueType', () => {
		expect(parseTrueType(new Uint8Array(40))).toBeUndefined();
		expect(
			parseTrueType(Uint8Array.from([0x4f, 0x54, 0x54, 0x4f, ...Array(30).fill(0)])),
		).toBeUndefined();
	});
});

describe('TrueType subsetting', () => {
	it('renumbers used glyphs from 1 and pulls in composite components', () => {
		const font = parseTrueType(syntheticFont())!;
		const subset = subsetTrueType(font, [2]);
		// B (composite) is glyph 1; its component A is appended as glyph 2.
		expect(subset.glyphs).toEqual([0, 2, 1]);
		const reread = parseTrueType(subset.data);
		expect(reread).toBeUndefined(); // the subset has no cmap by design
		const view = new DataView(subset.data.buffer);
		expect(view.getUint32(0)).toBe(0x00010000);
		const count = view.getUint16(4);
		const found = new Map<string, number>();
		for (let i = 0; i < count; i++) {
			const entry = 12 + i * 16;
			found.set(
				String.fromCharCode(...subset.data.subarray(entry, entry + 4)),
				view.getUint32(entry + 8),
			);
		}
		expect([...found.keys()]).toEqual(
			expect.arrayContaining(['glyf', 'loca', 'head', 'hmtx', 'maxp']),
		);
		// The composite's component id was rewritten from 1 to 2.
		const glyf = found.get('glyf')!;
		const loca = found.get('loca')!;
		const start = view.getUint32(loca + 4);
		expect(view.getUint16(glyf + start + 12)).toBe(2);
		expect(view.getUint16(found.get('maxp')! + 4)).toBe(3);
		// Whole-file checksum adjustment makes the file sum to 0xB1B0AFBA.
		let sum = 0;
		for (let i = 0; i < subset.data.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
		expect(sum).toBe(0xb1b0afba);
	});

	const arial = 'C:/Windows/Fonts/arial.ttf';
	it.skipIf(!existsSync(arial))('subsets a real font', () => {
		const font = parseTrueType(new Uint8Array(readFileSync(arial)))!;
		expect(font.glyphFor(0x2212)).toBeGreaterThan(0);
		const ids = [...'Hello é −'].map((c) => font.glyphFor(c.codePointAt(0)!));
		const subset = subsetTrueType(font, ids);
		expect(subset.data.length).toBeLessThan(font.data.length / 10);
		expect(subset.glyphs.length).toBeGreaterThan(6);
	});
});
