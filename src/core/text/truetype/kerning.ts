/**
 * Pair kerning of a TrueType font: the `kern` feature's pair-adjustment lookups in `GPOS`
 * (formats 1 and 2, through extension lookups too), else the legacy `kern` table's format 0
 * subtables. Scripts and languages are not told apart: every `kern` feature's lookups apply, which
 * is what the Latin, Greek and Cyrillic faces of the free Office substitutes have. DOM-free.
 */
import type { TrueTypeFont } from './parse';

/** The horizontal adjustment between two glyphs, in font units (negative pulls them together). */
export type PairKerning = (left: number, right: number) => number;

const u16 = (v: DataView, o: number) => (o >= 0 && o + 2 <= v.byteLength ? v.getUint16(o) : 0);
const i16 = (v: DataView, o: number) => (o >= 0 && o + 2 <= v.byteLength ? v.getInt16(o) : 0);
const u32 = (v: DataView, o: number) => (o >= 0 && o + 4 <= v.byteLength ? v.getUint32(o) : 0);

/** The coverage index of `glyph`, or -1. */
function coverage(view: DataView, at: number, glyph: number): number {
	const format = u16(view, at);
	const count = u16(view, at + 2);
	let lo = 0;
	let hi = count - 1;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		if (format === 1) {
			const value = u16(view, at + 4 + mid * 2);
			if (value === glyph) return mid;
			if (value < glyph) lo = mid + 1;
			else hi = mid - 1;
		} else if (format === 2) {
			const record = at + 4 + mid * 6;
			if (glyph < u16(view, record)) hi = mid - 1;
			else if (glyph > u16(view, record + 2)) lo = mid + 1;
			else return u16(view, record + 4) + glyph - u16(view, record);
		} else return -1;
	}
	return -1;
}

function classOf(view: DataView, at: number, glyph: number): number {
	if (!at) return 0;
	const format = u16(view, at);
	if (format === 1) {
		const start = u16(view, at + 2);
		const count = u16(view, at + 4);
		return glyph >= start && glyph < start + count ? u16(view, at + 6 + (glyph - start) * 2) : 0;
	}
	if (format !== 2) return 0;
	let lo = 0;
	let hi = u16(view, at + 2) - 1;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		const record = at + 4 + mid * 6;
		if (glyph < u16(view, record)) hi = mid - 1;
		else if (glyph > u16(view, record + 2)) lo = mid + 1;
		else return u16(view, record + 4);
	}
	return 0;
}

const bits = (value: number) => {
	let count = 0;
	for (let rest = value & 0xff; rest; rest &= rest - 1) count++;
	return count;
};
/** The X advance of a value record (0 when its format has none). */
const xAdvance = (view: DataView, record: number, format: number) =>
	format & 4 ? i16(view, record + bits(format & 3) * 2) : 0;

/** One pair-adjustment subtable: `undefined` when it does not cover the pair. */
function pairSubtable(view: DataView, at: number, left: number, right: number): number | undefined {
	const format = u16(view, at);
	const index = coverage(view, at + u16(view, at + 2), left);
	if (index < 0) return undefined;
	const format1 = u16(view, at + 4);
	const size = (bits(format1) + bits(u16(view, at + 6))) * 2;
	if (format === 1) {
		if (index >= u16(view, at + 8)) return undefined;
		const set = at + u16(view, at + 10 + index * 2);
		const record = 2 + size;
		let lo = 0;
		let hi = u16(view, set) - 1;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			const entry = set + 2 + mid * record;
			const second = u16(view, entry);
			if (second === right) return xAdvance(view, entry + 2, format1);
			if (second < right) lo = mid + 1;
			else hi = mid - 1;
		}
		return undefined;
	}
	if (format !== 2) return undefined;
	const class1 = classOf(view, at + u16(view, at + 8), left);
	const class2 = classOf(view, at + u16(view, at + 10), right);
	const class1Count = u16(view, at + 12);
	const class2Count = u16(view, at + 14);
	if (class1 >= class1Count || class2 >= class2Count) return undefined;
	return xAdvance(view, at + 16 + (class1 * class2Count + class2) * size, format1);
}

/** The pair-adjustment subtables of every `kern` feature lookup, in lookup order. */
function gposSubtables(view: DataView, gpos: number): number[] {
	const features = gpos + u16(view, gpos + 6);
	const lookups = gpos + u16(view, gpos + 8);
	const indices = new Set<number>();
	const featureCount = u16(view, features);
	for (let index = 0; index < featureCount; index++) {
		const record = features + 2 + index * 6;
		const tag = String.fromCharCode(
			view.getUint8(record),
			view.getUint8(record + 1),
			view.getUint8(record + 2),
			view.getUint8(record + 3),
		);
		if (tag !== 'kern') continue;
		const feature = features + u16(view, record + 4);
		const count = u16(view, feature + 2);
		for (let at = 0; at < count; at++) indices.add(u16(view, feature + 4 + at * 2));
	}
	const subtables: number[] = [];
	for (const lookupIndex of [...indices].sort((a, b) => a - b)) {
		if (lookupIndex >= u16(view, lookups)) continue;
		const lookup = lookups + u16(view, lookups + 2 + lookupIndex * 2);
		const type = u16(view, lookup);
		const count = u16(view, lookup + 4);
		for (let index = 0; index < count; index++) {
			const subtable = lookup + u16(view, lookup + 6 + index * 2);
			if (type === 2) subtables.push(subtable);
			else if (type === 9 && u16(view, subtable + 2) === 2)
				subtables.push(subtable + u32(view, subtable + 4));
		}
	}
	return subtables;
}

/** The format 0 pairs of a legacy (Microsoft) `kern` table, keyed `left << 16 | right`. */
function kernTablePairs(view: DataView, kern: number): Map<number, number> {
	const pairs = new Map<number, number>();
	if (u16(view, kern) !== 0) return pairs;
	let at = kern + 4;
	for (let table = u16(view, kern + 2); table > 0; table--) {
		const length = u16(view, at + 2);
		const cover = u16(view, at + 4);
		// Format 0, horizontal, not cross-stream, not minimum values.
		if (cover >> 8 === 0 && (cover & 7) === 1) {
			const count = u16(view, at + 6);
			for (let index = 0; index < count; index++) {
				const entry = at + 14 + index * 6;
				const key = u16(view, entry) * 65536 + u16(view, entry + 2);
				if (!pairs.has(key)) pairs.set(key, i16(view, entry + 4));
			}
		}
		if (!length) break;
		at += length;
	}
	return pairs;
}

/** The font's pair kerning, or `undefined` when it has neither a `GPOS` kern feature nor `kern`. */
export function pairKerning(font: TrueTypeFont): PairKerning | undefined {
	const view = new DataView(font.data.buffer, font.data.byteOffset, font.data.byteLength);
	const gpos = font.tables.get('GPOS');
	const subtables = gpos && gpos.length >= 10 ? gposSubtables(view, gpos.offset) : [];
	if (subtables.length) {
		const cache = new Map<number, number>();
		return (left, right) => {
			const key = left * 65536 + right;
			let value = cache.get(key);
			if (value === undefined) {
				value = 0;
				for (const subtable of subtables) {
					const found = pairSubtable(view, subtable, left, right);
					if (found !== undefined) {
						value = found;
						break;
					}
				}
				cache.set(key, value);
			}
			return value;
		};
	}
	const kern = font.tables.get('kern');
	if (!kern) return undefined;
	const pairs = kernTablePairs(view, kern.offset);
	return pairs.size ? (left, right) => pairs.get(left * 65536 + right) ?? 0 : undefined;
}
