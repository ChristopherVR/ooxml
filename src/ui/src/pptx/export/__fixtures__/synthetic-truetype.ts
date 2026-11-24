/**
 * Test fixture: a minimal TrueType font (glyphs for space, `A` and `B`), so PDF tests can check
 * font embedding without a system font. Built the same way as the core TrueType tests' font.
 */
const be16 = (v: number) => [(v >> 8) & 255, v & 255];
const be32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];

/** `extra` maps more characters (code point, glyph 1 to 3; in the BMP) to the glyphs. */
export function syntheticTrueType(extra: readonly (readonly [number, number])[] = []): Uint8Array {
	const simple = [...be16(1), ...be16(0), ...be16(0), ...be16(100), ...be16(100), 0, 0, 0, 0, 0, 0];
	const pad = (b: number[]) => (b.length % 4 ? [...b, ...Array(4 - (b.length % 4)).fill(0)] : b);
	const glyphs = [[], pad(simple), pad(simple), []] as number[][];
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
	const segments: (readonly [number, number, number])[] = [
		[0x20, 0x20, 3 - 0x20],
		[0x41, 0x42, 1 - 0x41],
		...extra.map(([code, glyph]) => [code, code, glyph - code] as const),
		[0xffff, 0xffff, 1],
	];
	segments.sort((a, b) => a[0] - b[0]);
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
