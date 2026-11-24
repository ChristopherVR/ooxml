/**
 * The tagged-PDF structure of the Word PDF export, as Word's "Document structure tags for
 * accessibility" writes it: a `Document` of headings (`H1`-`H6`), paragraphs (`P`), tables
 * (`Table`, `TR`, `TD`), figures with their alternative text and links, in reading order. Each
 * painted element is wrapped in its own marked-content sequence; borders, shading, headers,
 * footers and repeated table headings are artifacts. Disabled, every call is a no-op. DOM-free.
 */
import type { PdfStructElement, PdfVectorPage } from '../../export/pdf-vector';

type Open = PdfStructElement | 'artifact' | undefined;

export class DocxPdfTagger {
	readonly root: PdfStructElement = { type: 'Document', kids: [] };
	#page: PdfVectorPage | undefined;
	#open: Open;
	#suspended = 0;
	readonly #artifacts: boolean[] = [];
	readonly #containers: PdfStructElement[] = [this.root];
	readonly #paragraphs = new Map<string, PdfStructElement>();
	readonly #tables = new Map<string, PdfStructElement>();
	readonly #figures: { blockId: string; element: PdfStructElement }[] = [];

	constructor(readonly enabled: boolean) {}

	/** Whether content is tagged now (not off, and not inside an artifact). */
	get active(): boolean {
		return this.enabled && !this.#suspended && !!this.#page;
	}

	startPage(page: PdfVectorPage): void {
		this.#page = page;
		this.#open = undefined;
	}

	/** Closes the page's open sequence and files its floating figures under their paragraphs. */
	endPage(): void {
		this.close();
		for (const { blockId, element } of this.#figures.splice(0))
			(this.#paragraphs.get(blockId) ?? this.root).kids.push(element);
	}

	/** What is painted next belongs to `element`. */
	mark(element: PdfStructElement): void {
		if (!this.active || this.#open === element) return;
		this.close();
		element.kids.push({ page: this.#page!, mcid: this.#page!.beginMarked(element.type) });
		this.#open = element;
	}

	/** What is painted next is decoration (borders, shading, rules). */
	artifact(): void {
		if (!this.active || this.#open === 'artifact') return;
		this.close();
		this.#page!.beginArtifact();
		this.#open = 'artifact';
	}

	/** Ends the open sequence, before a clip (`q`) or its end (`Q`). */
	close(): void {
		if (!this.active || this.#open === undefined) return;
		this.#page!.endMarked();
		this.#open = undefined;
	}

	/** Inside a header, footer or repeated row: nothing is tagged or linked. */
	get inArtifact(): boolean {
		return this.#suspended > 0;
	}

	/** Paints `paint` as one pagination artifact (a header, footer or repeated row). */
	asArtifact(paint: () => void, subtype?: 'Header' | 'Footer'): void {
		this.suspend(subtype);
		try {
			paint();
		} finally {
			this.resume();
		}
	}

	/** Starts a pagination artifact, ended by `resume`. */
	suspend(subtype?: 'Header' | 'Footer'): void {
		const emit = this.active;
		if (emit) {
			this.close();
			this.#page!.beginArtifact('Pagination', subtype);
		}
		this.#suspended++;
		this.#artifacts.push(emit);
	}

	resume(): void {
		this.#suspended--;
		if (this.#artifacts.pop()) this.#page!.endMarked();
	}

	/** The element of a paragraph box; a box that continues one from an earlier page reuses it. */
	paragraph(blockId: string, type: string, continued: boolean): PdfStructElement | undefined {
		if (!this.active) return undefined;
		const existing = this.#paragraphs.get(blockId);
		if (existing && continued) return existing;
		const element: PdfStructElement = { type, kids: [] };
		this.#containers.at(-1)!.kids.push(element);
		this.#paragraphs.set(blockId, element);
		return element;
	}

	/** A child of `parent` (a `Link` or `Figure`) whose content follows. */
	child(parent: PdfStructElement, type: string, alt?: string): PdfStructElement | undefined {
		if (!this.active) return undefined;
		const element: PdfStructElement = { type, kids: [], ...(alt ? { alt } : {}) };
		parent.kids.push(element);
		return element;
	}

	/** A floating figure, filed under its anchor paragraph when the page ends. */
	figure(blockId: string, alt: string | undefined): PdfStructElement | undefined {
		if (!this.active) return undefined;
		const element: PdfStructElement = { type: 'Figure', kids: [], ...(alt ? { alt } : {}) };
		this.#figures.push({ blockId, element });
		return element;
	}

	/** Opens a table (continued across pages under one element), row or cell. */
	open(type: 'Table' | 'TR' | 'TD', blockId?: string): void {
		if (!this.active) return;
		this.close();
		let element = type === 'Table' && blockId ? this.#tables.get(blockId) : undefined;
		if (!element) {
			element = { type, kids: [] };
			this.#containers.at(-1)!.kids.push(element);
			if (type === 'Table' && blockId) this.#tables.set(blockId, element);
		}
		this.#containers.push(element);
	}

	end(): void {
		if (!this.active) return;
		this.close();
		if (this.#containers.length > 1) this.#containers.pop();
	}
}
