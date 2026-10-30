/**
 * A reader for TrueType-flavoured sfnt fonts (`.ttf`, and the first font of a `.ttc`): the
 * metrics, the character map and the outline tables a subsetter needs. OpenType fonts with
 * CFF outlines are refused (`parseTrueType` returns `undefined`). DOM-free and dependency free.
 */

export interface TrueTypeFont {
	readonly data: Uint8Array;
	readonly tables: ReadonlyMap<string, { offset: number; length: number }>;
	readonly unitsPerEm: number;
	readonly glyphCount: number;
	/** `hhea` ascent and descent in font units (descent is negative). */
	readonly ascent: number;
	readonly descent: number;
	readonly bbox: readonly [number, number, number, number];
	readonly italicAngle: number;
	/** Cap height in font units, when the `OS/2` table has it. */
	readonly capHeight: number | undefined;
	/** `OS/2` `fsType`: embedding permissions (0 = installable). */
	readonly embeddingFlags: number;
	/** 0 short, 1 long `loca` offsets. */
	readonly locaFormat: number;
	/** The glyph for a code point, or 0 when the font has none. */
	glyphFor(codePoint: number): number;
	/** Advance width of a glyph in font units. */
	advance(glyph: number): number;
	/** Left side bearing of a glyph in font units. */
	bearing(glyph: number): number;
	/** The bytes of a glyph's outline in `glyf` (empty for blank glyphs). */
	outline(glyph: number): Uint8Array;
}

const u16 = (v: DataView, o: number) => (o + 2 <= v.byteLength ? v.getUint16(o) : 0);
const i16 = (v: DataView, o: number) => (o + 2 <= v.byteLength ? v.getInt16(o) : 0);
const u32 = (v: DataView, o: number) => (o + 4 <= v.byteLength ? v.getUint32(o) : 0);
const tag = (v: DataView, o: number) =>
	String.fromCharCode(v.getUint8(o), v.getUint8(o + 1), v.getUint8(o + 2), v.getUint8(o + 3));

type Lookup = (codePoint: number) => number;

function format4(view: DataView, at: number): Lookup {
	const segments = u16(view, at + 6) / 2;
	const ends = at + 14;
	const starts = ends + segments * 2 + 2;
	const deltas = starts + segments * 2;
	const offsets = deltas + segments * 2;
	return (code) => {
		if (code > 0xffff) return 0;
		let lo = 0;
		let hi = segments - 1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			if (code > u16(view, ends + mid * 2)) lo = mid + 1;
			else if (code < u16(view, starts + mid * 2)) hi = mid - 1;
			else {
				const range = u16(view, offsets + mid * 2);
				if (!range) return (code + u16(view, deltas + mid * 2)) & 0xffff;
				const slot = offsets + mid * 2 + range + (code - u16(view, starts + mid * 2)) * 2;
				const glyph = u16(view, slot);
				return glyph ? (glyph + u16(view, deltas + mid * 2)) & 0xffff : 0;
			}
		}
		return 0;
	};
}

function format12(view: DataView, at: number): Lookup {
	const groups = u32(view, at + 12);
	return (code) => {
		let lo = 0;
		let hi = groups - 1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			const base = at + 16 + mid * 12;
			if (code > u32(view, base + 4)) lo = mid + 1;
			else if (code < u32(view, base)) hi = mid - 1;
			else return u32(view, base + 8) + (code - u32(view, base));
		}
		return 0;
	};
}

function characterMap(view: DataView, table: number): Lookup | undefined {
	const count = u16(view, table + 2);
	let best: { rank: number; at: number } | undefined;
	for (let index = 0; index < count; index++) {
		const platform = u16(view, table + 4 + index * 8);
		const encoding = u16(view, table + 6 + index * 8);
		const at = table + u32(view, table + 8 + index * 8);
		const format = u16(view, at);
		if (format !== 4 && format !== 12) continue;
		const rank =
			(platform === 3 && encoding === 10) || (platform === 0 && format === 12)
				? 3
				: platform === 3 && encoding === 1
					? 2
					: platform === 0
						? 1
						: 0;
		if (rank > 0 && (!best || rank > best.rank)) best = { rank, at };
	}
	if (!best) return undefined;
	return u16(view, best.at) === 12 ? format12(view, best.at) : format4(view, best.at);
}

/** Parse a TrueType font; `undefined` when it is not TrueType outlines or is malformed. */
export function parseTrueType(data: Uint8Array): TrueTypeFont | undefined {
	if (data.length < 12) return undefined;
	const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
	const base = tag(view, 0) === 'ttcf' ? u32(view, 12) : 0;
	if (base + 12 > data.length) return undefined;
	if (u32(view, base) !== 0x00010000 && tag(view, base) !== 'true') return undefined;
	const count = u16(view, base + 4);
	const tables = new Map<string, { offset: number; length: number }>();
	for (let index = 0; index < count; index++) {
		const entry = base + 12 + index * 16;
		if (entry + 16 > data.length) return undefined;
		const offset = u32(view, entry + 8);
		const length = u32(view, entry + 12);
		if (offset + length > data.length) return undefined;
		tables.set(tag(view, entry), { offset, length });
	}
	const head = tables.get('head');
	const hhea = tables.get('hhea');
	const maxp = tables.get('maxp');
	const hmtx = tables.get('hmtx');
	const loca = tables.get('loca');
	const glyf = tables.get('glyf');
	const cmap = tables.get('cmap');
	if (!head || !hhea || !maxp || !hmtx || !loca || !glyf || !cmap) return undefined;
	if (head.length < 54 || hhea.length < 36 || maxp.length < 6) return undefined;
	const unitsPerEm = u16(view, head.offset + 18);
	const glyphCount = u16(view, maxp.offset + 4);
	if (!unitsPerEm || !glyphCount) return undefined;
	const locaFormat = i16(view, head.offset + 50);
	const lookup = characterMap(view, cmap.offset);
	if (!lookup) return undefined;
	const metrics = u16(view, hhea.offset + 34);
	const os2 = tables.get('OS/2');
	const post = tables.get('post');
	return {
		data,
		tables,
		unitsPerEm,
		glyphCount,
		ascent: i16(view, hhea.offset + 4),
		descent: i16(view, hhea.offset + 6),
		bbox: [
			i16(view, head.offset + 36),
			i16(view, head.offset + 38),
			i16(view, head.offset + 40),
			i16(view, head.offset + 42),
		],
		italicAngle: post ? i16(view, post.offset + 4) + u16(view, post.offset + 6) / 65536 : 0,
		capHeight: os2 && os2.length >= 90 ? i16(view, os2.offset + 88) || undefined : undefined,
		embeddingFlags: os2 ? u16(view, os2.offset + 8) : 0,
		locaFormat,
		glyphFor: (code) => {
			const glyph = lookup(code);
			return glyph < glyphCount ? glyph : 0;
		},
		advance: (glyph) => u16(view, hmtx.offset + Math.min(glyph, Math.max(0, metrics - 1)) * 4),
		bearing: (glyph) =>
			glyph < metrics
				? i16(view, hmtx.offset + glyph * 4 + 2)
				: i16(view, hmtx.offset + metrics * 4 + (glyph - metrics) * 2),
		outline: (glyph) => {
			if (glyph < 0 || glyph >= glyphCount) return new Uint8Array(0);
			const at = loca.offset + glyph * (locaFormat ? 4 : 2);
			const start = locaFormat ? u32(view, at) : u16(view, at) * 2;
			const end = locaFormat ? u32(view, at + 4) : u16(view, at + 2) * 2;
			if (end <= start || glyf.offset + end > data.length) return new Uint8Array(0);
			return data.subarray(glyf.offset + start, glyf.offset + end);
		},
	};
}
