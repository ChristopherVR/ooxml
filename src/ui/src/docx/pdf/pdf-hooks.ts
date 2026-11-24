/**
 * What the block painter tells the export about each paragraph, line fragment and table it
 * draws, so the export can tag it (`pdf-tagging.ts`), link it (`pdf-navigation.ts`) and show
 * its tracked changes (`pdf-markup.ts`). Coordinates are CSS pixels, y down. DOM-free.
 */
import type { Block, Comment, DocumentModel, ReviewDisplayMode } from 'ooxml-core/docx';
import type {
	LayoutFloatBox,
	LayoutFragment,
	LayoutLine,
	LayoutPageBox,
	LayoutParagraphBox,
	LayoutTableRowBox,
} from 'ooxml-core/docx/layout';
import { pdfDate } from '../../export/pdf-archive';
import type { PdfStructElement, PdfVectorPage } from '../../export/pdf-vector';
import type { DocxPdfBalloons } from './pdf-balloon-items';
import type { DocxPdfBalloonFrame } from './pdf-balloons';
import { documentListItems, type DocxPdfListItem } from './pdf-lists';
import {
	commentNote,
	DocxPdfChangeBars,
	isChanged,
	isFormatChanged,
	markupFragment,
} from './pdf-markup';
import {
	DocxPdfNavigation,
	DocxPdfParagraphs,
	outlineLevel,
	runLinkTarget,
	type DocxPdfBookmarks,
	type DocxPdfLinkLabel,
	type DocxPdfLinkTarget,
} from './pdf-navigation';
import { DocxPdfTagger } from './pdf-tagging';

const PT_PER_PX = 0.75;

/** A list item's generated marker (its number or bullet and the separator after it). */
const isMarker = (box: LayoutParagraphBox, fragment: LayoutFragment, line: LayoutLine) =>
	line === box.lines[0] &&
	fragment.runIndex === 0 &&
	(fragment.sourceEnd ?? 0) <= (fragment.sourceStart ?? 0);

export interface DocxPdfHookOptions {
	tagged: boolean;
	bookmarks: DocxPdfBookmarks;
	mode: ReviewDisplayMode;
	/** The markup pane's balloons (`pdf-balloon-items.ts`), with markup shown. */
	balloons?: DocxPdfBalloons;
}

export class DocxPdfHooks {
	readonly tagger: DocxPdfTagger;
	readonly navigation: DocxPdfNavigation;
	readonly paragraphs: DocxPdfParagraphs;
	readonly #bars = new DocxPdfChangeBars();
	readonly #comments: Map<string, Comment>;
	readonly #noted = new Set<string>();
	#page: PdfVectorPage | undefined;
	#layout: LayoutPageBox | undefined;
	/** The page's shrink to make room for the markup pane, if any. */
	#frame: DocxPdfBalloonFrame | undefined;
	#boxStart = { x: 0, top: 0, height: 0 };
	#paragraph: PdfStructElement | undefined;
	/** The `Lbl` of the list item being painted, until its marker is drawn. */
	#marker: PdfStructElement | undefined;
	readonly #listItems: Map<string, DocxPdfListItem>;
	#link:
		| {
				target: DocxPdfLinkTarget;
				element: PdfStructElement | undefined;
				label: DocxPdfLinkLabel;
		  }
		| undefined;
	#decorative = 0;
	readonly #rows: boolean[] = [];

	constructor(
		readonly model: DocumentModel,
		readonly options: DocxPdfHookOptions,
	) {
		this.tagger = new DocxPdfTagger(options.tagged);
		this.navigation = new DocxPdfNavigation(model, options.bookmarks, options.mode);
		this.paragraphs = new DocxPdfParagraphs(model);
		this.#listItems = options.tagged ? documentListItems(model) : new Map();
		this.#comments = new Map((model.comments ?? []).map((comment) => [comment.id, comment]));
	}

	get markup(): boolean {
		return this.options.mode === 'all';
	}

	startPage(page: PdfVectorPage, layout: LayoutPageBox): void {
		this.#page = page;
		this.#layout = layout;
		this.#frame = this.options.balloons?.frame(layout);
		this.tagger.startPage(page);
	}

	endPage(): void {
		const page = this.#page;
		const layout = this.#layout;
		if (page && layout && this.#bars.size) {
			this.tagger.artifact();
			this.#bars.paint(page, layout.marginLeftPx);
		}
		if (page && layout && this.#frame) {
			this.tagger.artifact();
			this.options.balloons?.endPage(page, layout);
		}
		this.tagger.endPage();
	}

	/** A point of the page (CSS pixels, y down) in PDF points (y up), on the shrunk page if any. */
	#point(x: number, y: number): [number, number] {
		const scale = this.#frame?.scale ?? 1;
		return [
			x * PT_PER_PX * scale,
			(this.#frame?.offsetYPt ?? 0) + ((this.#layout?.heightPx ?? 0) - y) * PT_PER_PX * scale,
		];
	}

	/** Borders, shading and rules: not part of the text. */
	decoration(): void {
		this.tagger.artifact();
	}

	close(): void {
		this.tagger.close();
	}

	beginParagraph(box: LayoutParagraphBox, x: number, y: number): void {
		this.#link = undefined;
		this.#marker = undefined;
		const paragraph = this.paragraphs.get(box.blockId);
		const continued = (box.lines[0]?.sourceStart ?? 0) > 0;
		this.#boxStart = { x, top: y, height: box.lines[0]?.heightPx ?? 0 };
		this.options.balloons?.beginParagraph(paragraph, this.tagger.inArtifact);
		if (this.tagger.active) {
			const level = paragraph ? outlineLevel(paragraph, this.model.paragraphStyles) : undefined;
			// A numbered heading stays a heading, its number part of its text, as Word tags it.
			const item = level === undefined ? this.#listItems.get(box.blockId) : undefined;
			const listed = item && this.tagger.listItem(box.blockId, item, continued);
			if (listed) {
				this.#paragraph = listed.body;
				this.#marker = listed.label;
			} else {
				const type = level === undefined ? 'P' : `H${Math.min(6, level)}`;
				this.#paragraph = this.tagger.paragraph(box.blockId, type, continued);
			}
		} else this.#paragraph = undefined;
		if (paragraph && !continued && this.#page && !this.tagger.inArtifact) {
			const [left, top] = this.#point(x, y);
			this.navigation.paragraphStart(paragraph, { page: this.#page, x: left, y: top });
		}
	}

	/** The fragment to draw, after tagging it, recording its link and styling its markup. */
	fragment(
		box: LayoutParagraphBox,
		fragment: LayoutFragment,
		x: number,
		lineTop: number,
		line: LayoutLine,
	): LayoutFragment {
		const run =
			fragment.sourceEnd !== undefined && fragment.sourceEnd > (fragment.sourceStart ?? 0)
				? this.paragraphs.runAt(box.blockId, fragment.sourceStart)
				: fragment.object
					? this.paragraphs.runAt(box.blockId, fragment.sourceStart)
					: undefined;
		const paragraph = this.#paragraph;
		// Links in headers, footers and repeated rows are clickable on every page they are drawn.
		const target = run && this.#page ? runLinkTarget(run) : undefined;
		if (this.#marker && isMarker(box, fragment, line)) {
			this.#link = undefined;
			this.tagger.mark(this.#marker);
		} else if (target) {
			// Consecutive runs to the same place are one link.
			if (!this.#link || JSON.stringify(this.#link.target) !== JSON.stringify(target)) {
				const element = this.tagger.inArtifact
					? this.tagger.artifactLink()
					: paragraph && this.tagger.child(paragraph, 'Link');
				this.#link = { target, element, label: { text: '' } };
			}
			if (this.#link.element) this.tagger.mark(this.#link.element);
			this.#link.label.text += fragment.text;
			if (fragment.widthPx > 0) {
				const [left, bottom] = this.#point(x, lineTop + line.heightPx);
				const [right, top] = this.#point(x + fragment.widthPx, lineTop);
				const { element, label } = this.#link;
				this.navigation.link(this.#page!, target, [left, bottom, right, top], element, label);
			}
		} else {
			this.#link = undefined;
			if (fragment.object?.noteSeparator) this.tagger.artifact();
			else if (fragment.object && paragraph) {
				const alt = run?.image?.altText ?? run?.image?.title;
				const figure = this.tagger.child(paragraph, 'Figure', alt);
				if (figure) this.tagger.mark(figure);
			} else if (paragraph) this.tagger.mark(paragraph);
		}
		if (!this.markup || !this.#page) return fragment;
		const formatted = isFormatChanged(run, this.paragraphs.get(box.blockId));
		if (isChanged(run) || formatted) this.#bars.add(lineTop, line.heightPx);
		const balloons = this.options.balloons;
		balloons?.fragment(run, fragment, x, lineTop, line.heightPx);
		this.#comment(run, lineTop);
		const marked = markupFragment(fragment, run);
		return balloons ? balloons.style(marked, run) : marked;
	}

	#comment(run: { commentIds?: string[] } | undefined, lineTop: number): void {
		const layout = this.#layout;
		for (const id of run?.commentIds ?? []) {
			const comment = this.#comments.get(id);
			if (!comment || this.#noted.has(id) || !layout || !this.#page) continue;
			this.#noted.add(id);
			const replies = [...this.#comments.values()].filter((reply) => reply.parentId === id);
			const text = [comment.text, ...replies.map((reply) => `${reply.author}: ${reply.text}`)];
			const [x, top] = this.#point(layout.widthPx - layout.marginRightPx + 8, lineTop);
			const date = comment.date ? new Date(comment.date) : undefined;
			const note = commentNote(
				{
					...comment,
					text: text.join('\n'),
					...(date && !Number.isNaN(date.getTime()) ? { date: pdfDate(date) } : {}),
				},
				x,
				top,
			);
			this.#page.annotate(note);
			this.options.balloons?.comment(id, note, layout);
			if (this.#paragraph) {
				const annot = this.tagger.child(this.#paragraph, 'Annot');
				annot?.kids.push({ page: this.#page, annotation: note });
			}
		}
	}

	/** A floating object: a figure, with the alternative text of its picture. */
	float(float: LayoutFloatBox, paint: () => void): void {
		if (this.tagger.active) {
			const anchored = (this.paragraphs.get(float.blockId)?.runs ?? []).filter(
				(run) => run.image?.anchored,
			);
			const image = anchored[float.index]?.image;
			const figure = this.tagger.figure(float.blockId, image?.altText ?? image?.title);
			if (figure) this.tagger.mark(figure);
		}
		paint();
		this.tagger.close();
	}

	/** The header or footer about to be painted (its blocks with the page's field results). */
	story(blocks: readonly Block[] | undefined): void {
		this.paragraphs.story(blocks);
	}

	endParagraph(): void {
		this.options.balloons?.endParagraph(this.#boxStart);
		this.#link = undefined;
		this.#marker = undefined;
		this.tagger.close();
	}

	/** A table; `withText` false is a decorative nested frame (its text is in its cell). */
	beginTable(blockId: string, withText: boolean): void {
		if (!withText || this.#decorative) {
			this.#decorative++;
			this.tagger.artifact();
			return;
		}
		this.tagger.open('Table', blockId);
	}

	endTable(): void {
		if (this.#decorative) this.#decorative--;
		else this.tagger.end();
	}

	beginRow(row: LayoutTableRowBox): void {
		if (this.#decorative) return;
		// A heading row repeated on a later page is pagination, as Word tags it.
		this.#rows.push(row.repeated);
		if (row.repeated) this.tagger.suspend();
		else this.tagger.open('TR');
	}

	endRow(): void {
		if (this.#decorative) return;
		if (this.#rows.pop()) this.tagger.resume();
		else this.tagger.end();
	}

	beginCell(): void {
		if (!this.#decorative) this.tagger.open('TD');
	}

	endCell(): void {
		if (!this.#decorative) this.tagger.end();
	}
}
