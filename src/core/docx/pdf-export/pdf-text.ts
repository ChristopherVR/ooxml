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
	ptToPx,
	type LayoutFontSpec,
	type LayoutFragment,
	type LayoutLine,
	type TextMeasurer,
} from '../layout/index';
import type { PdfFontSet } from '../../pdf/write/pdf-embedded-font';
import {
	PDF_SYMBOL_FONT,
	encodeWinAnsi,
	pdfBaseFont,
	pdfTextWidth,
} from '../../pdf/write/pdf-fonts';
import { isSymbolEncodedFont, symbolFontBytes, symbolTextToUnicode } from '../../text/symbol-fonts';
import { parseColor, type Rgba } from '../../pdf/write/pdf-svg-style';
import type { PdfVectorPage } from '../../pdf/write/pdf-vector';
import { paintContentControlTag } from './pdf-content-control-tag';
import type { DocxPdfReportBuilder } from './pdf-report';

const DEFAULT_FAMILY = 'Calibri';
const LEADERS: Record<string, string> = { dot: '.', hyphen: '-', middleDot: '\u00b7' };

/** A CSS colour (or bare `RRGGBB`) as 0..1 channels; `null` for transparent or unknown. */
export function cssRgba(color: string | undefined): Rgba | null {
	if (!color) return null;
	try {
		return parseColor(/^[\da-f]{6}$/i.test(color) ? `#${color}` : color);
	} catch {
		return null;
	}
}

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

export class DocxPdfText {
	constructor(
		readonly fonts: PdfFontSet,
		readonly measurer: TextMeasurer,
		readonly report: DocxPdfReportBuilder,
	) {}

	/** The font a fragment is drawn in, at its drawn size (scripts are smaller). */
	#font(fragment: LayoutFragment): LayoutFontSpec {
		const sizePt = (fragment.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (fragment.script ? 0.65 : 1);
		return {
			family: fragment.fontFamily ?? DEFAULT_FAMILY,
			sizePx: (sizePt * 96) / 72,
			...(fragment.fontFallback ? { fallback: fragment.fontFallback } : {}),
			...(fragment.bold ? { bold: true } : {}),
			...(fragment.italic ? { italic: true } : {}),
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
		const baseline = this.baseline(fragment, line, lineTop);
		const top = fragment.topPx !== undefined ? lineTop + fragment.topPx : lineTop;
		if (fragment.contentControlTag) paintContentControlTag(page, fragment, x, top);
		else if (fragment.background && fragment.widthPx > 0) {
			const height = fragment.boxHeightPx ?? line.heightPx;
			if (fillWith(page, fragment.background, 'transparent')) {
				page.rect(x, top, fragment.widthPx, height);
				page.op('f');
			}
		}
		const text = fragment.leader
			? this.#leaderText(fragment, font)
			: fragment.text.replace(/[\t\n\r]/g, ' ');
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

	#leaderText(fragment: LayoutFragment, font: LayoutFontSpec): string {
		const character = LEADERS[fragment.leader!];
		if (!character) return '';
		return character.repeat(Math.max(0, Math.floor(fragment.widthPx / (font.sizePx * 0.28))));
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
		const parts = [...text.matchAll(/(?:\p{Ll}\p{M}*)+|[^\p{Ll}]+/gu)].map(([part]) => {
			const upper = part.toUpperCase();
			const face = upper === part ? font : { ...font, sizePx: font.sizePx * 0.8 };
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
		const bold = !!font.bold;
		const italic = !!font.italic;
		const spacing = fragment.characterSpacingPx ?? 0;
		const length = [...text].length;
		const fit = (natural: number) => {
			const target = fragment.widthPx - spacing * length;
			if (!(natural > 0) || !(target > 0)) return 100;
			return Math.min(200, Math.max(50, (target / natural) * 100));
		};
		// A trailing space whose advance (with word spacing, all scaled by Tz) spans the gap.
		const space = gapPx > 0 && !/\s$/.test(text);
		const wordSpacing = (scale: number, spaceNatural: number) =>
			(gapPx * 100) / scale - spaceNatural - spacing;
		if (isSymbolEncodedFont(family) && !this.fonts.pick(family, bold, italic, text)) {
			// Without the symbol font's file: Symbol's codes are the standard Symbol font's, and other
			// symbol fonts' pictures are written as their Unicode equivalents where they have one.
			const bytes = symbolFontBytes(family, text);
			if (bytes !== undefined) {
				this.report.substituted(family, PDF_SYMBOL_FONT);
				const scale = fit(pdfTextWidth(bytes, PDF_SYMBOL_FONT) * sizePx);
				page.text(PDF_SYMBOL_FONT, sizePx, x, y, bytes, { character: spacing, scale });
				return;
			}
			text = symbolTextToUnicode(family, text) ?? text;
		}
		const ascii = encodeWinAnsi(text);
		const embedded =
			this.fonts.prefers(family) || ascii === undefined
				? this.fonts.pick(family, bold, italic, text)
				: undefined;
		if (embedded) {
			if (embedded.family.toLowerCase() === family.toLowerCase())
				this.report.embedded(embedded.family);
			else this.report.substituted(family, embedded.family);
			const scale = fit(embedded.width(text) * sizePx);
			page.embeddedText(embedded, sizePx, x, y, space ? `${text} ` : text, {
				character: spacing,
				scale,
				...(space ? { word: wordSpacing(scale, embedded.width(' ') * sizePx) } : {}),
			});
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
		const base = pdfBaseFont(substitute, bold, italic);
		// Reported by family (Helvetica, Times, Courier): the style follows the text.
		this.report.substituted(family, base.replace(/-.*$/, ''));
		const scale = fit(pdfTextWidth(written, base) * sizePx);
		page.text(base, sizePx, x, y, space ? `${written} ` : written, {
			character: spacing,
			scale,
			...(space ? { word: wordSpacing(scale, pdfTextWidth(' ', base) * sizePx) } : {}),
		});
	}
}
