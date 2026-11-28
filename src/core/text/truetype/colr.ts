/**
 * A reader for the colour layers of a TrueType font: the `COLR` table's version 0 base glyph and
 * layer records and the first `CPAL` palette. A colour glyph (an emoji of Segoe UI Emoji) is its
 * layer glyphs drawn in order, each in its palette colour; a palette index of 0xFFFF means the
 * text colour. Version 1 paint graphs are not read (a font that has them still carries version 0
 * records for most of its glyphs). DOM-free and dependency free.
 */
import type { TrueTypeFont } from './parse';

/** One layer of a colour glyph: a glyph and its colour, `undefined` for the text colour. */
export interface ColourLayer {
	readonly glyph: number;
	readonly color: readonly [number, number, number] | undefined;
}

/** The layers of a base glyph, in drawing order; `undefined` for a glyph with no colour layers. */
export type ColourLayers = (glyph: number) => readonly ColourLayer[] | undefined;

/** The colour layers of `font`, or `undefined` when it has no usable `COLR` and `CPAL` tables. */
export function colourLayers(font: TrueTypeFont): ColourLayers | undefined {
	const colr = font.tables.get('COLR');
	const cpal = font.tables.get('CPAL');
	if (!colr || !cpal || colr.length < 14 || cpal.length < 12) return undefined;
	const view = new DataView(font.data.buffer, font.data.byteOffset, font.data.byteLength);
	const inside = (at: number, size: number, table: { offset: number; length: number }) =>
		at >= 0 && at + size <= table.length && table.offset + at + size <= view.byteLength;
	const bases = view.getUint16(colr.offset + 2);
	const baseAt = view.getUint32(colr.offset + 4);
	const layerAt = view.getUint32(colr.offset + 8);
	const layers = view.getUint16(colr.offset + 12);
	const entries = view.getUint16(cpal.offset + 2);
	const palettes = view.getUint16(cpal.offset + 4);
	const colorAt = view.getUint32(cpal.offset + 8);
	if (!palettes || !inside(baseAt, bases * 6, colr) || !inside(layerAt, layers * 4, colr))
		return undefined;
	const first = view.getUint16(cpal.offset + 12);
	const colour = (index: number): ColourLayer['color'] | null => {
		if (index === 0xffff) return undefined;
		const at = colorAt + (first + index) * 4;
		if (index >= entries || !inside(at, 4, cpal)) return undefined;
		const o = cpal.offset + at;
		// BGRA; a fully transparent colour paints nothing.
		return view.getUint8(o + 3) === 0
			? null
			: [view.getUint8(o + 2) / 255, view.getUint8(o + 1) / 255, view.getUint8(o) / 255];
	};
	const cache = new Map<number, readonly ColourLayer[] | undefined>();
	return (glyph) => {
		if (cache.has(glyph)) return cache.get(glyph);
		let lo = 0;
		let hi = bases - 1;
		let found: readonly ColourLayer[] | undefined;
		while (lo <= hi) {
			const mid = (lo + hi) >> 1;
			const record = colr.offset + baseAt + mid * 6;
			const id = view.getUint16(record);
			if (glyph > id) lo = mid + 1;
			else if (glyph < id) hi = mid - 1;
			else {
				const start = view.getUint16(record + 2);
				const count = Math.min(view.getUint16(record + 4), layers - start);
				const list: ColourLayer[] = [];
				for (let index = 0; index < count; index++) {
					const layer = colr.offset + layerAt + (start + index) * 4;
					const paint = colour(view.getUint16(layer + 2));
					if (paint !== null) list.push({ glyph: view.getUint16(layer), color: paint });
				}
				found = list.length ? list : undefined;
				break;
			}
		}
		cache.set(glyph, found);
		return found;
	};
}
