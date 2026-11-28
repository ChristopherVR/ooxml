/**
 * Test fixture: the synthetic TrueType font with `COLR` and `CPAL` tables added, so colour glyph
 * tests need no system font. U+2600 (glyph 1) is a colour glyph of three layers: glyph 2 in red
 * (palette entry 0), glyph 3 in blue (entry 1) and glyph 2 in the text colour (0xFFFF); `A` is
 * not. Palette 0 only. Built from `syntheticTrueType`.
 */
import { parseTrueType } from '../parse';
import { writeSfnt } from '../sfnt-write';
import { syntheticTrueType } from './synthetic-truetype';

const be16 = (v: number) => [(v >> 8) & 255, v & 255];
const be32 = (v: number) => [(v >>> 24) & 255, (v >> 16) & 255, (v >> 8) & 255, v & 255];

export function colourTrueType(): Uint8Array {
	const base = syntheticTrueType([[0x2600, 1]]);
	const font = parseTrueType(base)!;
	const tables: [string, Uint8Array][] = [...font.tables].map(([name, { offset, length }]) => [
		name,
		base.slice(offset, offset + length),
	]);
	// COLR v0: one base glyph record (glyph 1, first layer 0, three layers) and three layer records.
	const layers = [
		[2, 0],
		[3, 1],
		[2, 0xffff],
	];
	const colr = [
		...be16(0),
		...be16(1),
		...be32(14),
		...be32(20),
		...be16(layers.length),
		...be16(1),
		...be16(0),
		...be16(layers.length),
		...layers.flatMap(([glyph, palette]) => [...be16(glyph!), ...be16(palette!)]),
	];
	// CPAL v0: one palette of two colours (BGRA): red, then blue.
	const cpal = [
		...be16(0),
		...be16(2),
		...be16(1),
		...be16(2),
		...be32(14),
		...be16(0),
		0,
		0,
		255,
		255,
		255,
		0,
		0,
		255,
	];
	tables.push(['COLR', Uint8Array.from(colr)], ['CPAL', Uint8Array.from(cpal)]);
	tables.sort(([a], [b]) => (a < b ? -1 : 1));
	return writeSfnt(tables);
}
