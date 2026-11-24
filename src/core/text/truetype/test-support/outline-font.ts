/**
 * Test fixture: a small TrueType font with real outlines, for tests that draw glyphs. `A` is a
 * 600 x 700 square of straight lines, `B` a ring of off-curve points only (implied on-curve
 * midpoints), `C` a composite of `A` moved 100 units right and scaled by a half, and space is
 * blank. 1000 units per em, ascent 800, descent -200, installable.
 */
const be16 = (v: number) => [(v >> 8) & 255, v & 255];
const be32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];
const pad = (b: number[]) => (b.length % 4 ? [...b, ...Array(4 - (b.length % 4)).fill(0)] : b);

/** A simple glyph of one contour; points are `[x, y, onCurve]`. */
function simpleGlyph(points: readonly (readonly [number, number, boolean])[]): number[] {
	const xs = points.map((p) => p[0]);
	const ys = points.map((p) => p[1]);
	const deltas = (values: number[]) => values.map((v, i) => v - (values[i - 1] ?? 0));
	return [
		...be16(1),
		...be16(Math.min(...xs)),
		...be16(Math.min(...ys)),
		...be16(Math.max(...xs)),
		...be16(Math.max(...ys)),
		...be16(points.length - 1),
		...be16(0),
		...points.map((p) => (p[2] ? 1 : 0)),
		...deltas(xs).flatMap((v) => be16(v & 0xffff)),
		...deltas(ys).flatMap((v) => be16(v & 0xffff)),
	];
}

export const OUTLINE_FONT_SQUARE: readonly (readonly [number, number, boolean])[] = [
	[0, 0, true],
	[0, 700, true],
	[600, 700, true],
	[600, 0, true],
];

export function outlineTrueType(): Uint8Array {
	const ring = simpleGlyph([
		[300, 0, false],
		[0, 350, false],
		[300, 700, false],
		[600, 350, false],
	]);
	// Composite: flags ARG_1_AND_2_ARE_WORDS | ARGS_ARE_XY_VALUES | WE_HAVE_A_SCALE, glyph 1.
	const composite = [
		...be16(0xffff),
		...be16(100),
		...be16(0),
		...be16(400),
		...be16(350),
		...be16(0x0001 | 0x0002 | 0x0008),
		...be16(1),
		...be16(100),
		...be16(0),
		...be16(0x2000),
	];
	const glyphs = [[], pad(simpleGlyph(OUTLINE_FONT_SQUARE)), pad(ring), pad(composite), []];
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
		...be16(600),
		...be16(700),
		...be16(0),
		...be16(8),
		...be16(2),
		...be16(1),
		...be16(0),
	];
	const hhea = [...be32(0x00010000), ...be16(800), ...be16(-200), ...Array(26).fill(0), ...be16(5)];
	const maxp = [...be32(0x00010000), ...be16(5), ...Array(26).fill(0)];
	const hmtx = [500, 700, 700, 700, 300].flatMap((w) => [...be16(w), ...be16(0)]);
	// cmap format 4: space -> 4, A..C -> 1..3, end 0xFFFF.
	const segments = [
		[0x20, 0x20, 4 - 0x20],
		[0x41, 0x43, 1 - 0x41],
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
		{ name: 'cmap', bytes: cmap },
		{ name: 'glyf', bytes: glyphs.flat() },
		{ name: 'head', bytes: head },
		{ name: 'hhea', bytes: hhea },
		{ name: 'hmtx', bytes: hmtx },
		{ name: 'loca', bytes: loca },
		{ name: 'maxp', bytes: maxp },
	];
	const header = [...be32(0x00010000), ...be16(tables.length), ...Array(6).fill(0)];
	let offset = 12 + tables.length * 16;
	const directory: number[] = [];
	const body: number[] = [];
	for (const t of tables) {
		directory.push(...[...t.name].map((c) => c.charCodeAt(0)), ...be32(0), ...be32(offset));
		directory.push(...be32(t.bytes.length));
		const padded = pad(t.bytes);
		body.push(...padded);
		offset += padded.length;
	}
	return Uint8Array.from([...header, ...directory, ...body]);
}
