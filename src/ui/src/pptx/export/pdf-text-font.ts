/**
 * Which font a run of PDF text is written in (a supplied, embedded font or a standard one). Kept
 * apart from the slide painter so the public helpers do not load the PDF writer.
 */
import type { PdfEmbeddedFont, PdfFontSet } from 'ooxml-core/pdf/write';
import { encodeWinAnsi, pdfBaseFont } from 'ooxml-core/pdf/write';

/** How one run of text is written. */
export type PdfTextFont =
	| { kind: 'embedded'; font: PdfEmbeddedFont; named: boolean }
	| { kind: 'standard'; baseFont: string; encoded: string };

/** The first family of a CSS `font-family` list, unquoted. */
export function primaryFamily(family: string): string {
	const first = family.split(',')[0] ?? '';
	return first.trim().replace(/^['"]|['"]$/g, '');
}

/**
 * The font a run is written in: a supplied font of the run's family, else a standard font when
 * the text is WinAnsi, else any supplied font that has every character; `undefined` when none.
 */
export function resolvePdfTextFont(
	text: string,
	family: string,
	bold: boolean,
	italic: boolean,
	fonts?: PdfFontSet,
): PdfTextFont | undefined {
	if (fonts?.has(family)) {
		const font = fonts.pick(family, bold, italic, text);
		if (font) return { kind: 'embedded', font, named: true };
	}
	const encoded = encodeWinAnsi(text);
	if (encoded !== undefined)
		return { kind: 'standard', baseFont: pdfBaseFont(family, bold, italic), encoded };
	const fallback = fonts?.pick(family, bold, italic, text);
	return fallback ? { kind: 'embedded', font: fallback, named: false } : undefined;
}

/** Whether some font (supplied or standard) can write `text` as visible PDF text. */
export function canWritePdfText(
	text: string,
	family: string,
	bold: boolean,
	italic: boolean,
	fonts?: PdfFontSet,
): boolean {
	return resolvePdfTextFont(text, family, bold, italic, fonts) !== undefined;
}
