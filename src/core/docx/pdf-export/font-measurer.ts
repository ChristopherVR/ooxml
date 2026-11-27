/**
 * A text measurer built from font files, for laying a document out without a canvas (in Node):
 * advance widths from `hmtx`, pair kerning from `GPOS`/`kern` only where Word kerns (a run's
 * "Kerning for fonts" at or above its threshold, `w:kern`), and the single line height Windows
 * gives Word (see `verticalEm`), which is the CSS normal line height the canvas measurer reads for
 * the same fonts on Windows. A family with no supplied file resolves through `DocxPdfFontFiles` (host mapping,
 * then the substitution table); one that resolves to no file is estimated (`families` says
 * which). DOM-free.
 */
import { createFakeMeasurer, type LayoutFontSpec, type TextMeasurer } from '../layout/index';
import type { PdfFontFile } from '../../pdf/write/pdf-embedded-font';
import { designTextWidthEm, officeFontVerticalMetrics, symbolFontGlyph } from '../../text/index';
import { pairKerning, type PairKerning } from '../../text/truetype/kerning';
import { DocxPdfFontFiles, type DocxPdfFontMatch, type FontFace } from './font-files';

export interface FontFileMeasurer extends TextMeasurer {
	readonly files: DocxPdfFontFiles;
	/**
	 * Every family measured so far, with the supplied file that measured it (`undefined` when it
	 * was estimated: no supplied file and no supplied substitute).
	 */
	readonly families: ReadonlyMap<string, DocxPdfFontMatch | undefined>;
}

export interface FontFileMeasurerOptions {
	/** Host mapping from a document family to a supplied family (`{ Calibri: 'Carlito' }`). */
	substitutions?: Readonly<Record<string, string>>;
	/** Measures a family that resolves to no file; Office's measured advances, else a formula. */
	estimate?: TextMeasurer;
}

/**
 * Estimates text in a family with no file: the advance widths measured from the Office fonts
 * (`text/font-metrics`) where the family is one of them, else the fake measurer's formula.
 */
export function createEstimatedMeasurer(): TextMeasurer {
	const fake = createFakeMeasurer();
	const vertical = (font: LayoutFontSpec) => officeFontVerticalMetrics(font.family);
	return {
		widthOf(text, font) {
			const em = designTextWidthEm(text.replace(/[\t\n]/g, ''), font.family, !!font.bold);
			return em === undefined ? fake.widthOf(text, font) : em * font.sizePx;
		},
		lineHeightOf(font) {
			const metrics = vertical(font);
			return metrics
				? (metrics.ascent + metrics.descent + (metrics.lineGap ?? 0)) * font.sizePx
				: fake.lineHeightOf(font);
		},
		ascentOf(font) {
			const metrics = vertical(font);
			return metrics
				? (metrics.ascent + (metrics.lineGap ?? 0) / 2) * font.sizePx
				: fake.lineHeightOf(font) * 0.8;
		},
	};
}

/**
 * A face's single line height and the ascent within it, in em: the typographic metrics when the
 * font sets `USE_TYPO_METRICS`, else `usWinAscent + usWinDescent` plus the part of the `hhea` line
 * gap they leave out, as Windows gives them to Word (measured: Noto Serif, which sets the flag,
 * advances 1.362 em per line in Word; Carlito and Liberation Serif, which do not, 1.221 and 1.15).
 * The leading is split above and below the text, as the canvas measurer places it.
 */
function verticalEm(face: FontFace): { height: number; ascent: number } {
	const { font } = face;
	const em = font.unitsPerEm;
	const typo = font.typoMetrics;
	if (typo) {
		const gap = Math.max(0, typo.lineGap);
		return {
			height: (typo.ascent - typo.descent + gap) / em,
			ascent: (typo.ascent + gap / 2) / em,
		};
	}
	const hhea = font.ascent - font.descent;
	if (font.winAscent === undefined || font.winDescent === undefined) {
		const gap = Math.max(0, font.lineGap);
		return { height: (hhea + gap) / em, ascent: (font.ascent + gap / 2) / em };
	}
	const win = font.winAscent + font.winDescent;
	const leading = Math.max(0, font.lineGap - (win - hhea));
	return { height: (win + leading) / em, ascent: (font.winAscent + leading / 2) / em };
}

const SMALL_CAPS_SCALE = 0.8;

/** See the module comment. */
export function createFontFileMeasurer(
	fonts: readonly PdfFontFile[] | DocxPdfFontFiles,
	options: FontFileMeasurerOptions = {},
): FontFileMeasurer {
	const files =
		fonts instanceof DocxPdfFontFiles ? fonts : new DocxPdfFontFiles(fonts, options.substitutions);
	const estimate = options.estimate ?? createEstimatedMeasurer();
	const families = new Map<string, DocxPdfFontMatch | undefined>();
	const kerning = new Map<FontFace, PairKerning | undefined>();
	const widths = new Map<string, number>();
	const faceOf = (font: LayoutFontSpec): FontFace | undefined => {
		const match = files.match(font.family, font.fallback);
		if (!families.has(font.family)) families.set(font.family, match);
		return match && files.face(match.family, !!font.bold, !!font.italic);
	};
	const kernOf = (face: FontFace) => {
		if (!kerning.has(face)) kerning.set(face, pairKerning(face.font));
		return kerning.get(face);
	};
	/** Width in em of `text` in `face`; characters the face lacks are estimated. */
	const faceWidth = (text: string, face: FontFace, spec: LayoutFontSpec, kern: boolean) => {
		const { font } = face;
		const pairs = kern ? kernOf(face) : undefined;
		let units = 0;
		let previous = 0;
		let missing = '';
		for (const char of text) {
			const code = char.codePointAt(0)!;
			if (code === 9 || code === 10 || code === 13) {
				previous = 0;
				continue;
			}
			const glyph = symbolFontGlyph(font, face.family, code === 0xa0 ? 32 : code);
			if (!glyph) {
				missing += char;
				previous = 0;
				continue;
			}
			units += font.advance(glyph);
			if (pairs && previous) units += pairs(previous, glyph);
			previous = glyph;
		}
		const em = units / font.unitsPerEm;
		return missing ? em + estimate.widthOf(missing, spec) / spec.sizePx : em;
	};
	const measure = (text: string, spec: LayoutFontSpec): number => {
		const face = faceOf(spec);
		if (!face) return estimate.widthOf(text, spec);
		const kern = spec.kerning === 'normal';
		if (!spec.smallCaps) return faceWidth(text, face, spec, kern) * spec.sizePx;
		let width = 0;
		for (const [part] of text.matchAll(/(?:\p{Ll}\p{M}*)+|[^\p{Ll}]+/gu)) {
			const upper = part.toUpperCase();
			const size = upper === part ? spec.sizePx : spec.sizePx * SMALL_CAPS_SCALE;
			width += faceWidth(upper, face, { ...spec, sizePx: size }, kern) * size;
		}
		return width;
	};
	const vertical = (spec: LayoutFontSpec) => {
		const face = faceOf(spec);
		return face ? verticalEm(face) : undefined;
	};
	return {
		files,
		families,
		widthOf(text, spec) {
			const cacheKey = `${spec.family}\u0000${spec.sizePx}\u0000${spec.bold ? 1 : 0}${spec.italic ? 1 : 0}${spec.kerning === 'normal' ? 1 : 0}${spec.smallCaps ? 1 : 0}\u0000${text}`;
			let width = widths.get(cacheKey);
			if (width === undefined) {
				width = measure(text, spec);
				if (widths.size >= 20000) widths.delete(widths.keys().next().value!);
				widths.set(cacheKey, width);
			}
			return width;
		},
		lineHeightOf(spec) {
			const metrics = vertical(spec);
			return metrics ? metrics.height * spec.sizePx : estimate.lineHeightOf(spec);
		},
		ascentOf(spec) {
			const metrics = vertical(spec);
			if (metrics) return metrics.ascent * spec.sizePx;
			return estimate.ascentOf?.(spec) ?? estimate.lineHeightOf(spec) * 0.8;
		},
	};
}
