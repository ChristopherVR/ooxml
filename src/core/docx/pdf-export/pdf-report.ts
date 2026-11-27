import { pdfALevel, type PdfALevel } from '../../pdf/write/pdf-archive';
import type { PdfStructElement } from '../../pdf/write/pdf-vector';

/** What a Word PDF export approximated or left out, as Word's own export would not have. */
export interface DocxPdfReport {
	/** Pages written. */
	pages: number;
	/** Font families written with a font file the host supplied, embedded as subsets. */
	embeddedFonts: string[];
	/** Families written with a standard PDF font instead (not embedded), with the font used. */
	substitutedFonts: { family: string; standardFont: string }[];
	/** Characters no available font could write, drawn as `?`. */
	missingCharacters: string[];
	/**
	 * How the layout measured each family, when it measured from the supplied font files (the
	 * default in Node when `fonts` is given and no `measurer`): with the family's own file
	 * (`supplied`), another family's file from `fontSubstitutions` (`mapping`) or the substitution
	 * table (`substitute`), or with estimated metrics because no file was supplied (`estimated`,
	 * `font` is `null`). Absent when the host gave a measurer.
	 */
	fontMeasurement?: DocxPdfFontMeasurement[];
	/** Objects left out of the file, by kind, with the page numbers (1-based) they were on. */
	omitted: { kind: DocxPdfOmittedKind; reason: string; pages: number[] }[];
	/** The layout's own approximations relative to Word, and the painter's. */
	notes: string[];
	/** Set when PDF/A was requested: whether the file meets PDF/A-2, at which level, and why not. */
	pdfA?: DocxPdfConformance;
	/** Set when PDF/UA-1 was requested: whether the file meets PDF/UA-1, and why not. */
	pdfUa?: DocxPdfUaConformance;
}

export interface DocxPdfFontMeasurement {
	family: string;
	/** The supplied family whose file measured it, or `null` when it was estimated. */
	font: string | null;
	source: 'supplied' | 'mapping' | 'substitute' | 'estimated';
}

export interface DocxPdfConformance {
	conformant: boolean;
	/**
	 * The level the file declares (`pdfaid:conformance`): `A` (PDF/A-2a) when requested and met,
	 * `U` (PDF/A-2u) when all its text is in embedded fonts with Unicode maps, else `B`.
	 */
	level: PdfALevel;
	issues: string[];
}

export interface DocxPdfUaConformance {
	conformant: boolean;
	issues: string[];
}

const STANDARD_FAMILIES = new Set(['Helvetica', 'Times', 'Courier']);

/** Counts `Figure` elements in the structure tree that lack non-empty alternative text. */
export function figuresMissingAltText(root: PdfStructElement | undefined): number {
	if (!root) return 0;
	let missing = 0;
	const walk = (element: PdfStructElement) => {
		if (element.type === 'Figure' && !element.alt?.trim()) missing++;
		for (const kid of element.kids) if ('type' in kid) walk(kid);
	};
	walk(root);
	return missing;
}

export interface PdfConformanceCheck {
	requestedLevel?: 'A' | 'U';
	tagged?: boolean;
	structure?: PdfStructElement;
	language?: string;
	title?: string;
}

/**
 * PDF/A-2 needs every font embedded: text written in a standard PDF font breaks it. With every
 * font embedded the file meets level U (validated with veraPDF, `scripts/pdfa-validate.mjs`).
 * When `requestedLevel` is `'A'`, level A additionally requires structure tags, a document
 * language and alternative text on every figure.
 */
export function pdfAConformance(
	report: DocxPdfReport,
	pages: readonly { fonts: ReadonlyMap<string, string> }[],
	check: PdfConformanceCheck = {},
): DocxPdfConformance {
	const issues: string[] = [];
	const base = pdfALevel(pages);
	if (base !== 'U') {
		const families = report.substitutedFonts
			.filter((font) => STANDARD_FAMILIES.has(font.standardFont))
			.map((font) => font.family);
		issues.push(
			`Text${families.length ? ` in ${families.join(', ')}` : ''} uses standard PDF fonts, which are not embedded; supply the font files to embed them.`,
		);
	}
	if (check.requestedLevel === 'A') {
		if (!check.tagged || !check.structure)
			issues.push('PDF/A-2a requires structure tags (`taggedPdf: true`).');
		if (!check.language?.trim())
			issues.push('PDF/A-2a requires a document language (`language` option).');
		const missingAlt = figuresMissingAltText(check.structure);
		if (missingAlt > 0)
			issues.push(
				`${missingAlt} figure${missingAlt === 1 ? '' : 's'} lack alternative text, which PDF/A-2a requires.`,
			);
		return {
			conformant: issues.length === 0,
			level: issues.length === 0 ? 'A' : base,
			issues,
		};
	}
	return { conformant: issues.length === 0, level: base, issues };
}

/**
 * Checks PDF/UA-1 (ISO 14289-1) conformance: structure tags, embedded fonts with Unicode maps,
 * a document title, a natural language, and alternative text on every figure.
 */
export function pdfUaConformance(
	report: DocxPdfReport,
	pages: readonly { fonts: ReadonlyMap<string, string> }[],
	check: PdfConformanceCheck,
): DocxPdfUaConformance {
	const issues: string[] = [];
	if (!check.tagged || !check.structure)
		issues.push('PDF/UA-1 requires structure tags (`taggedPdf: true`).');
	if (pdfALevel(pages) !== 'U') {
		const families = report.substitutedFonts
			.filter((font) => STANDARD_FAMILIES.has(font.standardFont))
			.map((font) => font.family);
		issues.push(
			`Text${families.length ? ` in ${families.join(', ')}` : ''} uses standard PDF fonts, which are not embedded; supply the font files to embed them.`,
		);
	}
	if (!check.title?.trim())
		issues.push(
			'PDF/UA-1 requires a document title (`properties.title` with `documentProperties: true`).',
		);
	if (!check.language?.trim())
		issues.push('PDF/UA-1 requires a document language (`language` option).');
	const missingAlt = figuresMissingAltText(check.structure);
	if (missingAlt > 0)
		issues.push(
			`${missingAlt} figure${missingAlt === 1 ? '' : 's'} lack alternative text, which PDF/UA-1 requires.`,
		);
	return { conformant: issues.length === 0, issues };
}

export type DocxPdfOmittedKind =
	| 'picture'
	| 'chart'
	| 'shape'
	| 'textBox'
	| 'watermark'
	| 'headerFooterFloat'
	| 'pageBorder';

/** Collects the report while the pages are painted. */
export class DocxPdfReportBuilder {
	readonly #embedded = new Set<string>();
	readonly #substituted = new Map<string, string>();
	readonly #missing = new Set<string>();
	readonly #omitted = new Map<
		string,
		{ kind: DocxPdfOmittedKind; reason: string; pages: Set<number> }
	>();
	readonly #notes = new Set<string>();
	#measurement: DocxPdfFontMeasurement[] | undefined;
	/** The page being painted (1-based), for omissions. */
	page = 1;

	embedded(family: string): void {
		this.#embedded.add(family);
	}
	substituted(family: string, standardFont: string): void {
		if (!this.#substituted.has(family)) this.#substituted.set(family, standardFont);
	}
	missing(character: string): void {
		this.#missing.add(character);
	}
	omitted(kind: DocxPdfOmittedKind, reason: string): void {
		const key = `${kind}\u0000${reason}`;
		const entry = this.#omitted.get(key) ?? { kind, reason, pages: new Set<number>() };
		entry.pages.add(this.page);
		this.#omitted.set(key, entry);
	}
	note(message: string): void {
		this.#notes.add(message);
	}
	/** Records how the layout measured each family (see `DocxPdfReport.fontMeasurement`). */
	measured(measurement: DocxPdfFontMeasurement[]): void {
		this.#measurement = [...measurement].sort((a, b) => a.family.localeCompare(b.family));
		const estimated = this.#measurement.filter((entry) => entry.source === 'estimated');
		if (estimated.length)
			this.note(
				`No font file was supplied for ${estimated.map((entry) => entry.family).join(', ')}: its text was measured with estimated metrics, so its line and page breaks can differ from Word's.`,
			);
	}

	build(pages: number): DocxPdfReport {
		return {
			pages,
			embeddedFonts: [...this.#embedded].sort(),
			substitutedFonts: [...this.#substituted]
				.map(([family, standardFont]) => ({ family, standardFont }))
				.sort((a, b) => a.family.localeCompare(b.family)),
			missingCharacters: [...this.#missing],
			...(this.#measurement ? { fontMeasurement: this.#measurement } : {}),
			omitted: [...this.#omitted.values()].map(({ kind, reason, pages: on }) => ({
				kind,
				reason,
				pages: [...on].sort((a, b) => a - b),
			})),
			notes: [...this.#notes],
		};
	}
}
