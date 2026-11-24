/**
 * Links and bookmarks for the Word PDF export, as Word's Save As PDF writes them: hyperlinks
 * (external ones as URI actions, internal ones to the place of their bookmark), cross-references
 * and table of contents entries made with `\h` stay clickable (in headers and footers too, on
 * every page they are drawn), each with its text as its description, and headings or Word
 * bookmarks become the PDF's outline. Layout fragments are mapped back to the model's runs
 * through their source offsets. DOM-free.
 */
import {
	headingLevel,
	isRunHiddenForReview,
	resolveParagraphFormatting,
	type Block,
	type DocumentModel,
	type Paragraph,
	type ParagraphStyleCatalog,
	type ReviewDisplayMode,
	type TextRun,
} from 'ooxml-core/docx';
import { reviewSourceLength } from 'ooxml-core/docx/layout';
import { hyperlinkPolicy } from 'ooxml-core/opc';
import type {
	PdfDestination,
	PdfLinkAnnotation,
	PdfOutlineItem,
	PdfStructElement,
	PdfVectorPage,
} from '../../export/pdf-vector';

/** Word's "Create bookmarks using": none, the headings, or the document's own bookmarks. */
export type DocxPdfBookmarks = 'none' | 'headings' | 'wordBookmarks';

export type DocxPdfLinkTarget = { uri: string } | { anchor: string };

const OPEN_SCHEMES = ['tel', 'ftp'];

const external = (href: string): DocxPdfLinkTarget | undefined =>
	hyperlinkPolicy(href, { additionalOpenSchemes: OPEN_SCHEMES }).open ? { uri: href } : undefined;

/** The arguments of a field instruction, with quoted ones unquoted. */
const fieldArguments = (instr: string): string[] =>
	[...instr.matchAll(/"([^"]*)"|(\S+)/g)].map((match) => match[1] ?? match[2] ?? '');

/** Where a HYPERLINK field, or a REF, PAGEREF or NOTEREF field with `\h`, goes. */
export function fieldLinkTarget(instr: string): DocxPdfLinkTarget | undefined {
	const [name = '', ...args] = fieldArguments(instr);
	const kind = name.toUpperCase();
	if (kind === 'HYPERLINK') {
		let href: string | undefined;
		let anchor: string | undefined;
		for (let index = 0; index < args.length; index++) {
			const arg = args[index]!;
			if (arg.toLowerCase() === '\\l') anchor = args[++index];
			else if (arg.startsWith('\\')) index += /^\\[ostm]$/i.test(arg) ? 1 : 0;
			else href ??= arg;
		}
		if (href) return external(anchor ? `${href}#${anchor}` : href);
		return anchor ? { anchor } : undefined;
	}
	if ((kind === 'REF' || kind === 'PAGEREF' || kind === 'NOTEREF') && args[0])
		return args.some((arg) => arg.toLowerCase() === '\\h') ? { anchor: args[0] } : undefined;
	return undefined;
}

/** Where a run links to: its hyperlink, else the field whose result it shows. */
export function runLinkTarget(run: TextRun): DocxPdfLinkTarget | undefined {
	const link = run.link;
	if (link?.href) return external(link.anchor ? `${link.href}#${link.anchor}` : link.href);
	if (link?.anchor) return { anchor: link.anchor };
	return run.field ? fieldLinkTarget(run.field.instr) : undefined;
}

type ParagraphIndex = Map<string, { paragraph: Paragraph; starts: number[] }>;

function indexBlocks(index: ParagraphIndex, blocks: readonly Block[]): void {
	const add = (paragraph: Paragraph) => {
		const starts: number[] = [];
		let offset = 0;
		for (const run of paragraph.runs) {
			starts.push(offset);
			offset += reviewSourceLength(run);
		}
		index.set(paragraph.id, { paragraph, starts });
	};
	for (const block of blocks)
		if (block.type === 'paragraph') add(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(add);
}

/** Every paragraph of the document by id, with its runs' source offsets. */
export class DocxPdfParagraphs {
	readonly #paragraphs: ParagraphIndex = new Map();
	#story: ParagraphIndex | undefined;
	constructor(model: DocumentModel) {
		const blocks = [
			model.blocks,
			...[...(model.footnotes ?? []), ...(model.endnotes ?? [])].map((note) => note.blocks),
		];
		for (const list of blocks) indexBlocks(this.#paragraphs, list);
	}
	/**
	 * The header or footer being painted (with that page's field results), whose paragraphs are
	 * found before the document's; none after it.
	 */
	story(blocks: readonly Block[] | undefined): void {
		this.#story = undefined;
		if (blocks) indexBlocks((this.#story = new Map()), blocks);
	}
	#entry(id: string) {
		return this.#story?.get(id) ?? this.#paragraphs.get(id);
	}
	get(id: string): Paragraph | undefined {
		return this.#entry(id)?.paragraph;
	}
	/** The run holding source offset `offset` of paragraph `id`. */
	runAt(id: string, offset: number | undefined): TextRun | undefined {
		const entry = this.#entry(id);
		if (!entry || offset === undefined) return undefined;
		const { paragraph, starts } = entry;
		let low = 0;
		let high = starts.length - 1;
		while (low < high) {
			const middle = (low + high + 1) >> 1;
			if (starts[middle]! <= offset) low = middle;
			else high = middle - 1;
		}
		return paragraph.runs[low];
	}
}

/** The text a heading shows, as a bookmark title. */
function paragraphTitle(paragraph: Paragraph, mode: ReviewDisplayMode): string {
	return paragraph.runs
		.filter((run) => !run.fieldCode && !run.fieldChar && !run.noteReference && !run.image)
		.filter((run) => !run.vanish && !isRunHiddenForReview(run, mode))
		.map((run) => run.text.replace(/[\t\n\r]/g, ' '))
		.join('')
		.replace(/\s+/g, ' ')
		.trim();
}

/** The outline level (1-9) of a heading paragraph, or undefined for body text. */
export function outlineLevel(
	paragraph: Paragraph,
	catalog: ParagraphStyleCatalog | undefined,
): number | undefined {
	const level = catalog
		? resolveParagraphFormatting(paragraph, catalog).outlineLevel
		: paragraph.outlineLevel;
	if (level !== undefined) return level >= 1 && level <= 9 ? level : undefined;
	return headingLevel(paragraph, catalog);
}

interface PendingLink {
	page: PdfVectorPage;
	target: DocxPdfLinkTarget;
	rect: [number, number, number, number];
	element?: PdfStructElement;
	/** The text of the whole link (all its lines), its annotation's alternate description. */
	label?: DocxPdfLinkLabel;
}

/** The text of one link, gathered while its fragments are painted. */
export interface DocxPdfLinkLabel {
	text: string;
}

/** Collects destinations, links and bookmarks while the pages are painted. */
export class DocxPdfNavigation {
	readonly #destinations = new Map<string, PdfDestination>();
	readonly #headings: { level: number; item: PdfOutlineItem }[] = [];
	readonly #bookmarks: PdfOutlineItem[] = [];
	readonly #links: PendingLink[] = [];
	readonly #seen = new Set<string>();
	constructor(
		readonly model: DocumentModel,
		readonly bookmarks: DocxPdfBookmarks,
		readonly mode: ReviewDisplayMode,
	) {}

	/** Records where a paragraph starts: its bookmarks and, if it is a heading, its outline entry. */
	paragraphStart(paragraph: Paragraph, dest: PdfDestination): void {
		if (this.#seen.has(paragraph.id)) return;
		this.#seen.add(paragraph.id);
		for (const name of paragraph.bookmarks ?? []) {
			if (this.#destinations.has(name)) continue;
			this.#destinations.set(name, dest);
			if (this.bookmarks === 'wordBookmarks' && !name.startsWith('_'))
				this.#bookmarks.push({ title: name, dest });
		}
		if (this.bookmarks !== 'headings') return;
		const level = outlineLevel(paragraph, this.model.paragraphStyles);
		const title = level === undefined ? '' : paragraphTitle(paragraph, this.mode);
		if (level !== undefined && title) this.#headings.push({ level, item: { title, dest } });
	}

	/** A linked area of one line; it joins the previous one when they continue each other. */
	link(
		page: PdfVectorPage,
		target: DocxPdfLinkTarget,
		rect: [number, number, number, number],
		element?: PdfStructElement,
		label?: DocxPdfLinkLabel,
	): void {
		const last = this.#links.at(-1);
		if (
			last &&
			last.page === page &&
			last.element === element &&
			last.label === label &&
			JSON.stringify(last.target) === JSON.stringify(target) &&
			Math.abs(last.rect[1] - rect[1]) < 0.01 &&
			Math.abs(last.rect[3] - rect[3]) < 0.01 &&
			rect[0] >= last.rect[0] - 0.01 &&
			rect[0] <= last.rect[2] + 2
		) {
			last.rect[2] = Math.max(last.rect[2], rect[2]);
			return;
		}
		this.#links.push({
			page,
			target,
			rect,
			...(element ? { element } : {}),
			...(label ? { label } : {}),
		});
	}

	/** Places the link annotations; returns the outline and the anchors no page holds. */
	finish(): { outline: PdfOutlineItem[]; unresolved: string[] } {
		const unresolved = new Set<string>();
		for (const { page, target, rect, element, label } of this.#links) {
			let annotation: PdfLinkAnnotation | undefined;
			if ('uri' in target) annotation = { kind: 'link', rect, uri: target.uri };
			else {
				const dest = this.#destinations.get(target.anchor);
				if (dest) annotation = { kind: 'link', rect, dest };
				else unresolved.add(target.anchor);
			}
			if (!annotation) continue;
			// PDF/UA asks every link annotation for an alternate description (`/Contents`).
			const contents = label?.text.replace(/\s+/g, ' ').trim();
			if (contents) annotation.contents = contents;
			page.annotate(annotation);
			element?.kids.push({ page, annotation });
		}
		// A link whose target is in no page written is plain text in the structure.
		for (const { element } of this.#links)
			if (element && !element.kids.some((kid) => 'annotation' in kid)) element.type = 'Span';
		return { outline: this.#outline(), unresolved: [...unresolved] };
	}

	#outline(): PdfOutlineItem[] {
		if (this.bookmarks === 'wordBookmarks') return this.#bookmarks;
		const top: PdfOutlineItem[] = [];
		const stack: { level: number; item: PdfOutlineItem }[] = [];
		for (const entry of this.#headings) {
			while (stack.length && stack.at(-1)!.level >= entry.level) stack.pop();
			const parent = stack.at(-1)?.item;
			if (parent) (parent.children ??= []).push(entry.item);
			else top.push(entry.item);
			stack.push(entry);
		}
		return top;
	}
}
