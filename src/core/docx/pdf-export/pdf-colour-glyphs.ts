/**
 * Colour glyphs (Segoe UI Emoji) in the PDF export, loaded only for a document whose supplied
 * fonts include one with `COLR` and `CPAL` tables. Word writes such a glyph as its layers, each
 * layer glyph shown at the glyph's origin in its palette colour (palette 0; the text colour for
 * index 0xFFFF), and never the base glyph itself (measured on Word 16's PDF of three emoji).
 */
import type { PdfEmbeddedFont } from '../../pdf/write/pdf-embedded-font';
import type { PdfTextSpacing, PdfVectorPage } from '../../pdf/write/pdf-vector';
import { cssRgba } from './pdf-css-color';
import { colourLayers, type ColourLayers } from '../../text/truetype/colr';

const ZERO_WIDTH_SPACE = String.fromCharCode(0x200b);
const layersOf = new WeakMap<PdfEmbeddedFont, ColourLayers | undefined>();

/** The colour layers of a supplied font, read once. */
function layers(font: PdfEmbeddedFont): ColourLayers | undefined {
	if (!layersOf.has(font)) layersOf.set(font, colourLayers(font.font));
	return layersOf.get(font);
}

/**
 * Draws `text` with each colour glyph as its layers and the rest as ordinary text, advancing as
 * `PdfVectorPage.embeddedText` does, and returns true; false (nothing drawn) when no character of
 * `text` is a colour glyph. `color` is the text colour (a CSS colour), which the page's fill
 * colour is left as.
 */
export function drawColourText(
	page: PdfVectorPage,
	font: PdfEmbeddedFont,
	sizePx: number,
	x: number,
	y: number,
	text: string,
	spacing: PdfTextSpacing,
	color: string | undefined,
): boolean {
	const of = layers(font);
	const chars = [...text];
	const ink = cssRgba(color) ?? cssRgba('#000000');
	if (!of || !ink || !chars.some((char) => of(font.glyphOf(char.codePointAt(0)!)))) return false;
	const fill = [ink.r, ink.g, ink.b] as const;
	const scale = (spacing.scale ?? 100) / 100;
	const advance = (char: string, index: number) =>
		(font.width(char) * sizePx +
			(spacing.character ?? 0) +
			(char === ' ' ? (spacing.word ?? 0) : 0) +
			((spacing.kerning?.[index] ?? 0) * sizePx) / 1000) *
		scale;
	let start = 0;
	let segment = x;
	let cursor = x;
	const flush = (end: number) => {
		if (end > start)
			page.embeddedText(font, sizePx, segment, y, chars.slice(start, end).join(''), {
				...spacing,
				...(spacing.kerning ? { kerning: spacing.kerning.slice(start, end) } : {}),
			});
	};
	chars.forEach((char, index) => {
		const parts = of(font.glyphOf(char.codePointAt(0)!));
		if (parts) {
			flush(index);
			parts.forEach((layer, order) => {
				page.fillColor(...(layer.color ?? fill));
				// Only the first layer stands for the character when the text is extracted.
				page.embeddedText(font, sizePx, cursor, y, order ? ZERO_WIDTH_SPACE : char, {
					glyph: layer.glyph,
					...(spacing.skew ? { skew: spacing.skew } : {}),
					...(spacing.drawSize ? { drawSize: spacing.drawSize } : {}),
				});
			});
			page.fillColor(...fill);
			start = index + 1;
			segment = cursor + advance(char, index);
		}
		cursor += advance(char, index);
	});
	flush(chars.length);
	return true;
}
