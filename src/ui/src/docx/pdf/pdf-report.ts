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
	/** Objects left out of the file, by kind, with the page numbers (1-based) they were on. */
	omitted: { kind: DocxPdfOmittedKind; reason: string; pages: number[] }[];
	/** The layout's own approximations relative to Word, and the painter's. */
	notes: string[];
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

	build(pages: number): DocxPdfReport {
		return {
			pages,
			embeddedFonts: [...this.#embedded].sort(),
			substitutedFonts: [...this.#substituted]
				.map(([family, standardFont]) => ({ family, standardFont }))
				.sort((a, b) => a.family.localeCompare(b.family)),
			missingCharacters: [...this.#missing],
			omitted: [...this.#omitted.values()].map(({ kind, reason, pages: on }) => ({
				kind,
				reason,
				pages: [...on].sort((a, b) => a - b),
			})),
			notes: [...this.#notes],
		};
	}
}
