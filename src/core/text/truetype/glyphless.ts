/**
 * A TrueType font whose glyphs are all empty: glyph `i` has advance `advances[i]` (in 1/1000 em)
 * and no outline. A PDF uses it for an invisible text layer (render mode 3) over a picture of
 * text that no real font can write: the characters stay selectable and searchable through the
 * PDF font's ToUnicode map, and nothing is drawn. It holds the six tables an embedded TrueType
 * font in a PDF needs (`glyf`, `head`, `hhea`, `hmtx`, `loca`, `maxp`). DOM-free.
 */
import { writeSfnt } from './sfnt-write';

/** The most glyphs a TrueType font addresses. */
export const MAX_GLYPHLESS_GLYPHS = 0xffff;

/** A font of `advances.length` empty glyphs (at least one, `.notdef`), 1000 units per em. */
export function glyphlessTrueType(advances: readonly number[]): Uint8Array {
	const count = Math.min(MAX_GLYPHLESS_GLYPHS, Math.max(1, advances.length));
	const advance = (glyph: number) =>
		Math.max(0, Math.min(0xffff, Math.round(advances[glyph] ?? 0)));
	let widest = 0;
	const hmtx = new Uint8Array(count * 4);
	const hmtxView = new DataView(hmtx.buffer);
	for (let glyph = 0; glyph < count; glyph++) {
		hmtxView.setUint16(glyph * 4, advance(glyph));
		widest = Math.max(widest, advance(glyph));
	}
	const head = new Uint8Array(54);
	const headView = new DataView(head.buffer);
	headView.setUint32(0, 0x00010000); // version
	headView.setUint32(4, 0x00010000); // fontRevision
	headView.setUint32(12, 0x5f0f3cf5); // magicNumber
	headView.setUint16(16, 0x000b); // flags: baseline and left side bearing at 0, integer ppem
	headView.setUint16(18, 1000); // unitsPerEm
	headView.setInt16(38, -200); // yMin
	headView.setInt16(40, 1000); // xMax
	headView.setInt16(42, 800); // yMax
	headView.setUint16(46, 8); // lowestRecPPEM
	headView.setInt16(48, 2); // fontDirectionHint
	// indexToLocFormat (offset 50) stays 0: short `loca` offsets.
	const hhea = new Uint8Array(36);
	const hheaView = new DataView(hhea.buffer);
	hheaView.setUint32(0, 0x00010000);
	hheaView.setInt16(4, 800); // ascender
	hheaView.setInt16(6, -200); // descender
	hheaView.setUint16(10, widest); // advanceWidthMax
	hheaView.setInt16(18, 1); // caretSlopeRise
	hheaView.setUint16(34, count); // numberOfHMetrics
	const maxp = new Uint8Array(32);
	const maxpView = new DataView(maxp.buffer);
	maxpView.setUint32(0, 0x00010000);
	maxpView.setUint16(4, count); // numGlyphs
	maxpView.setUint16(14, 1); // maxZones
	// Every glyph is empty: all `loca` offsets are 0 and `glyf` holds only padding.
	const loca = new Uint8Array((count + 1) * 2);
	const glyf = new Uint8Array(4);
	return writeSfnt([
		['glyf', glyf],
		['head', head],
		['hhea', hhea],
		['hmtx', hmtx],
		['loca', loca],
		['maxp', maxp],
	]);
}
