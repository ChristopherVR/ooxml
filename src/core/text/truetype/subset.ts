/**
 * Subset a TrueType font to a list of glyphs. The listed glyphs become glyph ids 1..n in the
 * order given (id 0 stays `.notdef`), so a PDF can address them by sequential CIDs; glyphs
 * that composite outlines reference are appended after them with their ids rewritten.
 * Hinting programs (`cvt `, `fpgm`, `prep`) are dropped: PDF viewers rasterise unhinted and the
 * tables are most of a small subset's size. DOM-free.
 */
import type { TrueTypeFont } from './parse';
import { pad4, writeSfnt } from './sfnt-write';

export interface TrueTypeSubset {
	/** The font file bytes. */
	readonly data: Uint8Array;
	/** Original glyph id per new glyph id (index 0 is `.notdef`). */
	readonly glyphs: readonly number[];
}

/** Component glyph ids of a composite outline, with the byte offset of each id field. */
function components(outline: Uint8Array): { offset: number; glyph: number }[] {
	if (outline.length < 10) return [];
	const view = new DataView(outline.buffer, outline.byteOffset, outline.byteLength);
	if (view.getInt16(0) >= 0) return [];
	const found: { offset: number; glyph: number }[] = [];
	let at = 10;
	for (let guard = 0; guard < 256 && at + 4 <= outline.length; guard++) {
		const flags = view.getUint16(at);
		found.push({ offset: at + 2, glyph: view.getUint16(at + 2) });
		if (!(flags & 0x20)) break;
		at += 4 + (flags & 1 ? 4 : 2);
		if (flags & 0x8) at += 2;
		else if (flags & 0x40) at += 4;
		else if (flags & 0x80) at += 8;
	}
	return found;
}

export function subsetTrueType(font: TrueTypeFont, used: readonly number[]): TrueTypeSubset {
	const glyphs: number[] = [0];
	const mapped = new Map<number, number>([[0, 0]]);
	const add = (glyph: number) => {
		if (mapped.has(glyph) || glyph >= font.glyphCount) return;
		mapped.set(glyph, glyphs.length);
		glyphs.push(glyph);
	};
	for (const glyph of used) add(glyph);
	// Close over composite components; the list grows while it is walked.
	for (let index = 1; index < glyphs.length; index++)
		for (const part of components(font.outline(glyphs[index]!))) add(part.glyph);

	const outlines = glyphs.map((glyph) => {
		const source = font.outline(glyph);
		if (!source.length) return source;
		const copy = Uint8Array.from(source);
		const view = new DataView(copy.buffer);
		for (const part of components(copy)) view.setUint16(part.offset, mapped.get(part.glyph) ?? 0);
		return copy;
	});

	const count = glyphs.length;
	let glyfLength = 0;
	for (const outline of outlines) glyfLength += pad4(outline.length);
	const glyf = new Uint8Array(glyfLength);
	const loca = new Uint8Array((count + 1) * 4);
	const locaView = new DataView(loca.buffer);
	let cursor = 0;
	outlines.forEach((outline, index) => {
		locaView.setUint32(index * 4, cursor);
		glyf.set(outline, cursor);
		cursor += pad4(outline.length);
	});
	locaView.setUint32(count * 4, cursor);

	const hmtx = new Uint8Array(count * 4);
	const hmtxView = new DataView(hmtx.buffer);
	glyphs.forEach((glyph, index) => {
		hmtxView.setUint16(index * 4, font.advance(glyph));
		hmtxView.setInt16(index * 4 + 2, font.bearing(glyph));
	});

	const copyTable = (name: string): Uint8Array | undefined => {
		const entry = font.tables.get(name);
		return entry
			? Uint8Array.from(font.data.subarray(entry.offset, entry.offset + entry.length))
			: undefined;
	};
	const head = copyTable('head')!;
	const hhea = copyTable('hhea')!;
	const maxp = copyTable('maxp')!;
	new DataView(head.buffer).setUint32(8, 0);
	new DataView(head.buffer).setInt16(50, 1);
	new DataView(hhea.buffer).setUint16(34, count);
	new DataView(maxp.buffer).setUint16(4, count);

	const tables: [string, Uint8Array][] = [];
	for (const [name, bytes] of [
		['glyf', glyf],
		['head', head],
		['hhea', hhea],
		['hmtx', hmtx],
		['loca', loca],
		['maxp', maxp],
	] as const)
		if (bytes) tables.push([name, bytes]);

	return { data: writeSfnt(tables), glyphs };
}
