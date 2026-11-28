/**
 * Text in the standard Symbol font that has codes with no glyph: each of those is drawn as the
 * font's missing-glyph box (an outlined rectangle), the rest as usual, all fitted to the width the
 * layout gave the fragment. Kept apart from `pdf-text.ts` and loaded with the symbol tables, since
 * only a symbol font's text without its file needs it. DOM-free.
 */
import type { LayoutFontSpec, LayoutFragment, TextMeasurer } from '../layout/index';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';
import type { SymbolMissingGlyph } from '../../text/symbol-fonts';

/** Draws `text` in pieces: `drawText(piece, left, widthPx, gapPx)` writes one that has a glyph. */
export function drawWithMissingGlyphs(
	page: PdfVectorPage,
	text: string,
	font: LayoutFontSpec,
	x: number,
	y: number,
	fragment: LayoutFragment,
	gapPx: number,
	missingOf: (family: string, char: string) => SymbolMissingGlyph | undefined,
	measurer: TextMeasurer,
	drawText: (piece: string, left: number, widthPx: number, gapPx: number) => void,
): void {
	const parts: { text: string; missing: SymbolMissingGlyph | undefined; width: number }[] = [];
	for (const char of text) {
		const missing = missingOf(font.family, char);
		const last = parts[parts.length - 1];
		if (last && last.missing === missing) last.text += char;
		else parts.push({ text: char, missing, width: 0 });
	}
	for (const part of parts) {
		part.width = part.missing
			? [...part.text].length * font.sizePx * part.missing.advance
			: measurer.widthOf(part.text, font);
	}
	const natural = parts.reduce((sum, part) => sum + part.width, 0);
	if (natural <= 0) return;
	let left = x;
	for (const [index, part] of parts.entries()) {
		const widthPx = (fragment.widthPx * part.width) / natural;
		if (part.missing) {
			const count = [...part.text].length;
			for (let at = 0; at < count; at++) {
				const shape = part.missing;
				// The font's own proportions, stretched with the fragment to the width it was given.
				const stretch = widthPx / count / (shape.advance * font.sizePx);
				const em = font.sizePx;
				const boxLeft = left + at * (widthPx / count) + shape.left * em * stretch;
				const boxWidth = shape.width * em * stretch;
				const height = shape.height * em;
				const wall = shape.wall * em;
				const wallX = wall * stretch;
				page.rect(boxLeft, y - height, boxWidth, height);
				page.rect(boxLeft + wallX, y - height + wall, boxWidth - 2 * wallX, height - 2 * wall);
				page.op('f*');
			}
		} else {
			drawText(part.text, left, widthPx, index === parts.length - 1 ? gapPx : 0);
		}
		left += widthPx;
	}
}
