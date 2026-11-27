/**
 * What a vector PDF export approximated, as structured codes with parameters: no sentence is
 * baked in, so every binding shows the same notice in the viewer's language (the
 * `pptx.pdfExport.diagnostic.<code>` keys). The collector and painter add to a
 * {@link PdfDiagnosticSet}; the export result carries the list. Framework-free.
 */
import { PDF_SVG_NOTE_CODES, type PdfSvgNoteCode } from 'ooxml-core/pdf/write';

/** An element (or its effect) is drawn as a picture of just that element. */
export const PDF_PICTURE_CODES = [
	'filterAsPicture',
	'blendAsPicture',
	'threeDAsPicture',
	'mediaAsPicture',
	'textShadowAsPicture',
	'svgAsPicture',
	'fillTextAsPicture',
	'boxFillAsPicture',
	'borderAsPicture',
	'shadowAsPicture',
	'verticalTextAsPicture',
	'textColourAsPicture',
	'complexScriptAsPicture',
	'noFontTextAsPicture',
] as const;

/** A whole slide is one picture (its text kept searchable). */
export const PDF_SLIDE_CODES = [
	'slideTooMuchText',
	'slideTooManyEffects',
	'slideNotCaptured',
] as const;

/** Something could not be written at all and is missing from the file. */
export const PDF_OMITTED_CODES = [
	'backgroundPictureMissing',
	'pictureMissing',
	'effectMissing',
	'drawingPartMissing',
	'unsearchableText',
] as const;

/** A font the slide asks for is not the one in the file (params `family`, `substitute`). */
export const PDF_FONT_CODES = ['fontStandard', 'fontSubstituted'] as const;

/** Drawn as vector graphics, but simpler than on the slide. */
export const PDF_SIMPLIFIED_CODES = ['groupOpacity', ...PDF_SVG_NOTE_CODES] as const;

/** Information only: not an approximation by itself. */
export const PDF_NOTE_CODES = ['searchableText'] as const;

export type PdfExportDiagnosticCode =
	| (typeof PDF_PICTURE_CODES)[number]
	| (typeof PDF_SLIDE_CODES)[number]
	| (typeof PDF_OMITTED_CODES)[number]
	| (typeof PDF_FONT_CODES)[number]
	| (typeof PDF_NOTE_CODES)[number]
	| 'groupOpacity'
	| PdfSvgNoteCode;

export type PdfExportDiagnosticCategory =
	| 'picture'
	| 'slide'
	| 'omitted'
	| 'font'
	| 'simplified'
	| 'note';

/** One thing a vector PDF approximated. */
export interface PdfExportDiagnostic {
	readonly code: PdfExportDiagnosticCode;
	/** Interpolated into the message (`family`, `substitute`). */
	readonly params?: Readonly<Record<string, string>>;
}

const CATEGORY = new Map<PdfExportDiagnosticCode, PdfExportDiagnosticCategory>([
	...PDF_PICTURE_CODES.map((code) => [code, 'picture'] as const),
	...PDF_SLIDE_CODES.map((code) => [code, 'slide'] as const),
	...PDF_OMITTED_CODES.map((code) => [code, 'omitted'] as const),
	...PDF_FONT_CODES.map((code) => [code, 'font'] as const),
	...PDF_SIMPLIFIED_CODES.map((code) => [code, 'simplified'] as const),
	...PDF_NOTE_CODES.map((code) => [code, 'note'] as const),
]);

/** Every code, in the order the notice lists them. */
export const PDF_EXPORT_DIAGNOSTIC_CODES: readonly PdfExportDiagnosticCode[] = [...CATEGORY.keys()];

/** What kind of approximation `code` is. */
export function pdfDiagnosticCategory(code: PdfExportDiagnosticCode): PdfExportDiagnosticCategory {
	return CATEGORY.get(code) ?? 'note';
}

/** The dictionary key of the message for `code`. */
export function pdfDiagnosticMessageKey(code: PdfExportDiagnosticCode): string {
	return `pptx.pdfExport.diagnostic.${code}`;
}

const identity = (diagnostic: PdfExportDiagnostic): string =>
	diagnostic.params
		? `${diagnostic.code}|${Object.entries(diagnostic.params)
				.sort(([a], [b]) => a.localeCompare(b))
				.map(([key, value]) => `${key}=${value}`)
				.join('|')}`
		: diagnostic.code;

/** Distinct diagnostics in the order first seen (same code and params count once). */
export class PdfDiagnosticSet {
	readonly #items = new Map<string, PdfExportDiagnostic>();

	add(code: PdfExportDiagnosticCode, params?: Readonly<Record<string, string>>): void {
		const diagnostic: PdfExportDiagnostic = params ? { code, params } : { code };
		const key = identity(diagnostic);
		if (!this.#items.has(key)) this.#items.set(key, diagnostic);
	}

	addAll(diagnostics: Iterable<PdfExportDiagnostic>): void {
		for (const diagnostic of diagnostics) this.add(diagnostic.code, diagnostic.params);
	}

	get size(): number {
		return this.#items.size;
	}

	toArray(): PdfExportDiagnostic[] {
		return [...this.#items.values()];
	}
}

/** `diagnostics` without repeats (same code and params), first seen first. */
export function distinctPdfDiagnostics(
	diagnostics: Iterable<PdfExportDiagnostic>,
): PdfExportDiagnostic[] {
	const set = new PdfDiagnosticSet();
	set.addAll(diagnostics);
	return set.toArray();
}
