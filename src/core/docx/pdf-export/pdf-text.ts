/**
 * Writes the text fragments of a laid-out line as real PDF text: a host-supplied font embedded
 * as a subset when one matches, else a standard PDF font. Each fragment is scaled (`Tz`) to the
 * width the layout measured, so line breaks and positions match Print Layout whichever font
 * writes it. DOM-free.
 */
import {
	DEFAULT_FONT_SIZE_PT,
	contentControlTagMetrics,
	contentControlTagStyle,
	contentControlTagTitleOffset,
	fontMetrics,
	kerningMode,
	ptToPx,
	type LayoutFontSpec,
	type LayoutFragment,
	type LayoutLine,
	type TextMeasurer,
} from '../layout/index';
import type { PdfFontSet } from '../../pdf/write/pdf-embedded-font';
import type { PdfCidFont } from '../../pdf/write/pdf-cid-font';
import {
	PDF_SYMBOL_FONT,
	encodeWinAnsi,
	pdfBaseFont,
	pdfTextWidth,
} from '../../pdf/write/pdf-fonts';
import { isSymbolEncodedFont } from '../../text/index';

/** The symbol-font conversions the text writer uses, loaded with the PDF export's glyph tables. */
export type DocxPdfSymbolTables = Pick<
	typeof import('../../text/symbol-fonts'),
	'symbolFontBytes' | 'symbolTextToUnicode' | 'symbolMissingGlyph'
>;
import { cssRgba } from './pdf-css-color';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';
import { paintContentControlTag } from './pdf-content-control-tag';
import type { DocxPdfReportBuilder } from './pdf-report';
import { SMALL_CAPS_SCALE, smallCapsParts } from '../caps-display';
import { wordBaselinePx, wordPdfPx } from './pdf-units';

export { cssRgba };

const DEFAULT_FAMILY = 'Calibri';
/** The slope Word shears a face with when the font family has no italic (measured from its PDF). */
const SYNTHETIC_ITALIC_SKEW = 1 / 3;
/**
 * The pen width, in em of the drawn size, that emboldens a face the family has no bold of (render
 * mode 2 with the text colour as the pen). Measured on Word 16's PDFs of Sylfaen and Ink Free at
 * 8, 11, 16, 24 and 48 pt in two colours: the width is always the drawn size / 35 (0.22971,
 * 0.31543, 0.456, 0.68571 and 1.3714 pt), and the glyphs keep the regular face's advances.
 */
const SYNTHETIC_BOLD_EM = 1 / 35;
/**
 * Word draws a leader's characters on a grid of its advance anchored at the page's left edge (a
 * character starts at a whole multiple of its advance), in the cells that lie wholly between the
 * text before the tab and the text after it (issues #396 and #432, measured on Calibri 11 pt and
 * Arial 10 pt, with 1417 and 1500 twip margins). A cell may overhang the text after it by
 * this share of an advance, to absorb rounding.
 */
const LEADER_END_TOLERANCE = 0.01;
/** A cell may start this share of an advance before the text before it ends (a tab stop's rounding). */
const LEADER_START_TOLERANCE = 0.05;
/** Private-use characters: no font a document names covers them unless it is a symbol font. */
const PRIVATE_USE = /[-]/gu;
const LEADERS: Record<string, string> = { dot: '.', hyphen: '-', middleDot: '\u00b7' };

/** Sets the fill colour from a CSS colour; returns false for a transparent one. */
export function fillWith(
	page: PdfVectorPage,
	color: string | undefined,
	fallback = '#000000',
): boolean {
	const rgba = color === undefined ? cssRgba(fallback) : (cssRgba(color) ?? cssRgba(fallback));
	if (!rgba || rgba.a === 0) return false;
	page.fillColor(rgba.r, rgba.g, rgba.b);
	return true;
}

/** Pair kerning after each character of `text` in 1/1000 em (the last has none), if the font has any. */
function pairAdjustments(font: PdfCidFont, text: string): number[] | undefined {
	if (!font.kern) return undefined;
	const chars = [...text];
	return chars.map((char, index) => {
		const next = chars[index + 1];
		return next === undefined ? 0 : font.kern!(char, next) * 1000;
	});
}

export class DocxPdfText {
	/** Colour glyph drawing, set when a supplied font has `COLR` and `CPAL` tables. */
	colour: typeof import('./pdf-colour-glyphs') | undefined;
	/** Off while a watermark is drawn: Word draws those as outlines of the face, never emboldened. */
	emboldenMissingBold = true;
	constructor(
		readonly fonts: PdfFontSet,
		readonly measurer: TextMeasurer,
		readonly report: DocxPdfReportBuilder,
		readonly symbols?: DocxPdfSymbolTables,
		/** The missing-glyph box drawing, loaded with the symbol tables. */
		readonly boxes?: typeof import('./pdf-missing-glyph-boxes'),
	) {}

	/** The font a fragment is drawn in, at its drawn size (scripts are smaller). */
	#font(fragment: LayoutFragment): LayoutFontSpec {
		const kerning = kerningMode(fragment.kerningThresholdPt, fragment.fontSizePt);
		const sizePt = (fragment.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (fragment.script ? 0.65 : 1);
		return {
			family: fragment.fontFamily ?? DEFAULT_FAMILY,
			sizePx: (sizePt * 96) / 72,
			...(fragment.fontFallback ? { fallback: fragment.fontFallback } : {}),
			...(fragment.bold ? { bold: true } : {}),
			...(fragment.italic ? { italic: true } : {}),
			...(kerning ? { kerning } : {}),
		};
	}

	/**
	 * The baseline of `fragment` on a line whose top is at `lineTop`: the fragment box the layout
	 * placed (top and line height, as Print Layout draws a span), else the line's baseline.
	 */
	baseline(fragment: LayoutFragment, line: LayoutLine, lineTop: number): number {
		if (fragment.contentControlTag && line.baselinePx !== undefined)
			return lineTop + line.baselinePx - ptToPx(contentControlTagMetrics.labelRisePt);
		if (fragment.topPx !== undefined && fragment.boxHeightPx !== undefined) {
			const { ascent, descent } = fontMetrics(this.measurer, this.#font(fragment));
			return lineTop + fragment.topPx + (fragment.boxHeightPx - ascent - descent) / 2 + ascent;
		}
		return lineTop + (line.baselinePx ?? line.heightPx * 0.8);
	}

	/**
	 * Draws one text fragment (or tab leader) with its highlight and lines. `gapPx` is the room the
	 * layout left before the next word: it is written as a real space stretched to fill it, so the
	 * PDF's text copies and extracts with its spaces while every word keeps its position.
	 */
	draw(
		page: PdfVectorPage,
		fragment: LayoutFragment,
		x: number,
		line: LayoutLine,
		lineTop: number,
		gapPx = 0,
	): void {
		const font = this.#font(fragment);
		// Word puts baselines on whole 1/600 inch steps of the space the text is drawn in.
		const baseline = this.#wordBaseline(fragment, line, lineTop);
		const top = fragment.topPx !== undefined ? lineTop + fragment.topPx : lineTop;
		if (fragment.contentControlTag) paintContentControlTag(page, fragment, x, top);
		else if (fragment.background && fragment.widthPx > 0) {
			const height = fragment.boxHeightPx ?? line.heightPx;
			if (fillWith(page, fragment.background, 'transparent')) {
				page.rect(x, top, fragment.widthPx, height);
				page.op('f');
			}
		}
		// A leader's characters keep their own advance and end where the leader does, so the dots
		// of successive lines share one grid, as in Word.
		const leader = fragment.leader ? this.#leader(fragment, font, x) : undefined;
		if (leader) {
			x = leader.xPx;
			fragment = { ...fragment, widthPx: leader.widthPx };
		}
		const text = fragment.leader ? (leader?.text ?? '') : fragment.text.replace(/[\t\n\r]/g, ' ');
		if (fragment.leader === 'underscore' || fragment.leader === 'heavy') {
			this.#rule(
				page,
				fragment,
				x,
				baseline + font.sizePx * 0.1,
				fragment.leader === 'heavy' ? 2 : 1,
			);
		} else if (text.trim()) {
			const tag = fragment.contentControlTag;
			fillWith(page, tag ? contentControlTagStyle.text : fragment.color);
			if (tag) {
				const offset = ptToPx(contentControlTagTitleOffset(tag, !!fragment.blockContentControlTag));
				const widthPx = this.measurer.widthOf(text, font);
				this.#text(page, text, font, x + offset, baseline, { ...fragment, widthPx }, 0);
			} else if (fragment.smallCaps && !fragment.leader)
				this.#smallCaps(page, text, font, x, baseline, fragment, gapPx);
			else this.#text(page, text, font, x, baseline, fragment, fragment.leader ? 0 : gapPx);
		}
		const thickness = Math.max(0.75, font.sizePx * 0.06);
		if (fragment.underline && !fragment.leader)
			this.#rule(page, fragment, x, baseline + font.sizePx * 0.12, thickness);
		if (fragment.strike) this.#rule(page, fragment, x, baseline - font.sizePx * 0.28, thickness);
	}

	/**
	 * The baseline on Word's grid: Word's line rule (`wordBaselinePx`) for text on a single-spaced
	 * line's baseline, moved by any raise or lowering, else the baseline rounded to a whole row.
	 */
	#wordBaseline(fragment: LayoutFragment, line: LayoutLine, lineTop: number): number {
		const baseline = this.baseline(fragment, line, lineTop);
		if (line.baselinePx === undefined || line.extraBelowPx || fragment.contentControlTag)
			return wordPdfPx(baseline);
		const shift = baseline - (lineTop + line.baselinePx);
		const ruled = wordBaselinePx(lineTop, line.heightPx, line.baselinePx);
		return Math.abs(shift) < 1e-6 ? ruled : wordPdfPx(ruled + shift);
	}

	/**
	 * The leader's characters at their own advance, in the cells of the page-wide grid that fit
	 * between `x` (where the text before the tab ends) and the end of the leader's width.
	 */
	#leader(
		fragment: LayoutFragment,
		font: LayoutFontSpec,
		x: number,
	): { text: string; widthPx: number; xPx: number } | undefined {
		const character = LEADERS[fragment.leader!];
		if (!character) return undefined;
		// Word spaces the characters on its 1/600 inch grid: a 10 pt Arial period is 2.76 pt.
		const measured = wordPdfPx(this.measurer.widthOf(character, font));
		const advance = measured > 0 ? measured : font.sizePx * 0.28;
		const first = Math.ceil(x / advance - LEADER_START_TOLERANCE);
		const last = Math.floor((x + fragment.widthPx) / advance + LEADER_END_TOLERANCE);
		const count = Math.max(0, last - first);
		return { text: character.repeat(count), widthPx: count * advance, xPx: first * advance };
	}

	#rule(
		page: PdfVectorPage,
		fragment: LayoutFragment,
		x: number,
		y: number,
		thickness: number,
	): void {
		if (!fragment.widthPx || !fillWith(page, fragment.color)) return;
		page.rect(x, y - thickness / 2, fragment.widthPx, thickness);
		page.op('f');
	}

	#smallCaps(
		page: PdfVectorPage,
		text: string,
		font: LayoutFontSpec,
		x: number,
		y: number,
		fragment: LayoutFragment,
		gapPx: number,
	): void {
		const parts = smallCapsParts(text).map(({ text: upper, small }) => {
			const face = small ? { ...font, sizePx: font.sizePx * SMALL_CAPS_SCALE } : font;
			return { text: upper, face, width: this.measurer.widthOf(upper, face) };
		});
		const natural = parts.reduce((sum, part) => sum + part.width, 0);
		if (natural <= 0) return;
		for (const [index, part] of parts.entries()) {
			const widthPx = (fragment.widthPx * part.width) / natural;
			this.#text(
				page,
				part.text,
				part.face,
				x,
				y,
				{ ...fragment, widthPx },
				index === parts.length - 1 ? gapPx : 0,
			);
			x += widthPx;
		}
	}

	/**
	 * The run's own font for text with private-use characters it has no glyph for: Word draws such
	 * a character with the font's missing-glyph glyph (glyph 0) and keeps the rest of the word in
	 * the font (#430), where no font covering all of `text` would write it all as `?` in a standard
	 * font. Only the family's own file counts; a character of a script falls back to another font.
	 */
	#withNotdef(family: string, bold: boolean, italic: boolean, text: string) {
		const covered = text.replace(PRIVATE_USE, '');
		if (covered.length === text.length || !this.fonts.has(family)) return undefined;
		const font = this.fonts.pick(family, bold, italic, covered);
		return font?.family.toLowerCase() === family.toLowerCase() ? font : undefined;
	}

	#text(
		page: PdfVectorPage,
		text: string,
		font: LayoutFontSpec,
		x: number,
		y: number,
		fragment: LayoutFragment,
		gapPx: number,
	): void {
		const { family, sizePx } = font;
		text = text.replaceAll('\uFE0F', '');
		if (!text) return;
		// Word draws the glyphs at the size rounded to 1/600 inch, on the advances of the nominal one.
		const drawPx = wordPdfPx(sizePx);
		const bold = !!font.bold;
		const italic = !!font.italic;
		const spacing = fragment.characterSpacingPx ?? 0;
		const length = [...text].length;
		const fit = (natural: number) => {
			// The pair kerning with the next fragment's first glyph is part of the width, not of this text.
			const target = fragment.widthPx - (fragment.kernAfterPx ?? 0) - spacing * length;
			if (!(natural > 0) || !(target > 0)) return 100;
			return Math.min(200, Math.max(50, (target / natural) * 100));
		};
		// A trailing space whose advance (with word spacing, all scaled by Tz) spans the gap.
		const space = gapPx > 0 && !/\s$/.test(text);
		const wordSpacing = (scale: number, spaceNatural: number) =>
			(gapPx * 100) / scale - spaceNatural - spacing;
		// Without the symbol font's file: Symbol's codes are the standard Symbol font's, and other
		// symbol fonts' pictures are written as their Unicode equivalents where they have one. PDF/A
		// takes a supplied font that draws the equivalent over the unembedded Symbol font.
		let symbolBytes: string | undefined;
		if (
			this.symbols &&
			isSymbolEncodedFont(family) &&
			!this.fonts.pick(family, bold, italic, text)
		) {
			const unicode = this.symbols.symbolTextToUnicode(family, text);
			if (!this.fonts.embedAll || !unicode || !this.fonts.pick(family, bold, italic, unicode))
				symbolBytes = this.symbols.symbolFontBytes(family, text);
			text = unicode ?? text;
		}
		if (
			this.symbols &&
			isSymbolEncodedFont(family) &&
			[...text].some((char) => this.symbols!.symbolMissingGlyph(family, char)) &&
			!this.fonts.pick(family, bold, italic, text)
		) {
			this.boxes!.drawWithMissingGlyphs(
				page,
				text,
				font,
				x,
				y,
				fragment,
				gapPx,
				this.symbols.symbolMissingGlyph,
				this.measurer,
				(piece, left, widthPx, gap) =>
					this.#text(page, piece, font, left, y, { ...fragment, widthPx }, gap),
			);
			return;
		}
		const ascii = symbolBytes ?? encodeWinAnsi(text);
		const embedded =
			symbolBytes === undefined && (this.fonts.prefers(family) || ascii === undefined)
				? (this.fonts.pick(family, bold, italic, text) ??
					this.#withNotdef(family, bold, italic, text))
				: undefined;
		if (embedded) {
			if (embedded.family.toLowerCase() === family.toLowerCase())
				this.report.embedded(embedded.family);
			else this.report.substituted(family, embedded.family);
			// Kerned text is drawn pair by pair at the nominal widths (Word applies no scaling for it).
			const kerning = font.kerning === 'normal' ? pairAdjustments(embedded, text) : undefined;
			const kerned = kerning ? kerning.reduce((sum, value) => sum + value, 0) / 1000 : 0;
			const scale = fit((embedded.width(text) + kerned) * sizePx);
			// The family has no bold face: thicken the regular one, as Word does.
			const embolden = this.emboldenMissingBold && bold && !embedded.bold;
			if (embolden) {
				const pen = cssRgba(fragment.color) ?? cssRgba('#000000');
				page.save();
				if (pen) page.strokeColor(pen.r, pen.g, pen.b);
				page.lineStyle(drawPx * SYNTHETIC_BOLD_EM, 0, 0);
			}
			const options = {
				...(embolden ? { emboldened: true } : {}),
				character: spacing,
				...(kerning?.some(Boolean) ? { kerning } : {}),
				scale,
				// The family has no italic face: slant the regular one, as Word does.
				...(italic && !embedded.italic ? { skew: SYNTHETIC_ITALIC_SKEW } : {}),
				...(drawPx !== sizePx ? { drawSize: drawPx } : {}),
				...(space ? { word: wordSpacing(scale, embedded.width(' ') * sizePx) } : {}),
			};
			const shown = space ? `${text} ` : text;
			// A colour glyph (COLR/CPAL) is drawn as its coloured layers, never as the base outline.
			// A colour glyph (COLR/CPAL) is drawn as its coloured layers, never as the base outline.
			if (
				!this.colour?.drawColourText(page, embedded, sizePx, x, y, shown, options, fragment.color)
			)
				page.embeddedText(embedded, sizePx, x, y, shown, options);
			if (embolden) page.restore();
			return;
		}
		let written = ascii;
		if (written === undefined) {
			written = '';
			for (const character of text) {
				const code = encodeWinAnsi(character);
				if (code === undefined) this.report.missing(character);
				written += code ?? '?';
			}
		}
		const substitute =
			font.fallback?.alternate ??
			(font.fallback?.generic === 'serif'
				? 'Times New Roman'
				: font.fallback?.generic === 'monospace'
					? 'Courier New'
					: family);
		const base =
			symbolBytes === undefined ? pdfBaseFont(substitute, bold, italic) : PDF_SYMBOL_FONT;
		// Reported by family (Helvetica, Times, Courier): the style follows the text.
		this.report.substituted(family, base.replace(/-.*$/, ''));
		// A standard font is fitted to the layout's width with `Tz` anyway, so it is drawn at the
		// rounded size directly (its `Tj` has no per-glyph offsets).
		const scale = fit(pdfTextWidth(written, base) * drawPx);
		page.text(base, drawPx, x, y, space ? `${written} ` : written, {
			character: spacing,
			scale,
			...(space ? { word: wordSpacing(scale, pdfTextWidth(' ', base) * drawPx) } : {}),
		});
	}
}
