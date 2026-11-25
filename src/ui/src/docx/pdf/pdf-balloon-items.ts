/**
 * Collects the markup balloons of each page while its paragraphs are painted (`pdf-hooks.ts`)
 * and paints them in the markup pane when the page ends (`pdf-balloons.ts`). Which revisions go
 * in balloons follows Word's "Balloons" setting: comments and formatting by default, every
 * revision (deletions too, which then leave the text), only comments, or none (no pane: the page
 * keeps its size and comments are only PDF notes). DOM-free.
 */
import {
	describeCellFormatChange,
	describeParagraphFormatChange,
	describeRowFormatChange,
	describeRunFormatChange,
	describeSectionFormatChange,
	describeTableFormatChange,
	formatChangeText,
	listRevisions,
	type DocumentModel,
	type Paragraph,
	type Revision,
	type SectionProperties,
	type Table,
	type TextRun,
} from 'ooxml-core/docx';
import {
	reviewSourceLength,
	type LayoutFragment,
	type LayoutPageBox,
} from 'ooxml-core/docx/layout';
import { authorColor } from 'ooxml-core/docx/ui';
import type { PdfNoteAnnotation, PdfVectorPage } from '../../export/pdf-vector';
import { DocxPdfBalloonComments, paleColor } from './pdf-balloon-comments';
import {
	balloonFrame,
	LABEL_FAMILY,
	paintBalloons,
	type Balloon,
	type BalloonPiece,
	type DocxPdfBalloonFrame,
} from './pdf-balloons';
import type { DocxPdfReportBuilder } from './pdf-report';
import type { DocxPdfText } from './pdf-text';

/**
 * Word's "Balloons" setting for markup: `commentsAndFormatting` (Word's default) puts comments
 * and formatting changes in balloons and shows insertions and deletions in the text; `all` puts
 * deletions in balloons too; `comments` shows every revision in the text and only comments in
 * balloons (Word's "Show all revisions inline", which still prints comment balloons); `none`
 * draws no markup pane, leaving comments as PDF notes.
 */
export type DocxPdfMarkupBalloons = 'all' | 'commentsAndFormatting' | 'comments' | 'none';

const PT_PER_PX = 0.75;
const MOVE_COLOR = '#2e7d32';

interface Deletion {
	start: number;
	end: number;
	text: string;
	revision: Revision;
}

interface Anchor {
	x: number;
	top: number;
	height: number;
}

export class DocxPdfBalloons {
	readonly active: boolean;
	readonly #comments: DocxPdfBalloonComments;
	readonly #family: string;
	readonly #labelFamily: string;
	readonly #deletions = new Map<string, Deletion[]>();
	readonly #seen = new Set<string>();
	readonly #tables = new Map<string, Table>();
	readonly #sections = new Map<string, SectionProperties>();
	#items: Balloon[] = [];
	#pageComments = new Map<string, Balloon>();
	#paragraph: { paragraph: Paragraph; last: Anchor | undefined; lastEnd: number } | undefined;

	constructor(
		readonly model: DocumentModel,
		readonly mode: DocxPdfMarkupBalloons,
		markup: boolean,
		readonly text: DocxPdfText,
		readonly report: DocxPdfReportBuilder,
	) {
		this.#family =
			model.characterStyles?.docDefaults.fontFamily ?? model.theme?.fonts.minor.latin ?? 'Calibri';
		// Word labels balloons in Segoe UI; without that font file, the document's font keeps
		// every font embedded that the host supplied.
		this.#labelFamily = text.fonts.has(LABEL_FAMILY) ? LABEL_FAMILY : this.#family;
		this.#comments = new DocxPdfBalloonComments(model.comments, {
			text: this.#family,
			label: this.#labelFamily,
		});
		let hasStructuralFormatChange = false;
		for (const block of model.blocks) {
			if (block.type !== 'table') continue;
			this.#tables.set(block.id, block);
			if (block.formatRevision) hasStructuralFormatChange = true;
			if (block.rowProperties?.some((row) => row.formatRevision)) hasStructuralFormatChange = true;
			if (block.rows.some((row) => row.some((cell) => cell.formatRevision)))
				hasStructuralFormatChange = true;
		}
		for (const section of model.sections ?? []) {
			if (section.endsAtBlockId) this.#sections.set(section.endsAtBlockId, section);
			if (section.formatRevision) hasStructuralFormatChange = true;
		}
		const kinds = listRevisions(model).map((entry) => entry.kind);
		this.active =
			markup &&
			mode !== 'none' &&
			(this.#comments.size > 0 ||
				(this.formatting &&
					(hasStructuralFormatChange ||
						kinds.some((kind) => kind === 'formatChange' || kind === 'paragraphChange'))) ||
				(this.deletions && kinds.some((kind) => kind === 'delete' || kind === 'moveFrom')));
	}

	get formatting(): boolean {
		return this.mode === 'all' || this.mode === 'commentsAndFormatting';
	}

	/** Deletions leave the text for balloons (the layout must hide them). */
	get deletions(): boolean {
		return this.mode === 'all';
	}

	/** How a page is shrunk to make room for the pane; undefined without balloons. */
	frame(layout: LayoutPageBox): DocxPdfBalloonFrame | undefined {
		return this.active ? balloonFrame(layout) : undefined;
	}

	/**
	 * A paragraph box starts. Headers and footers (`inStory`) are laid out with their deletions
	 * in the text, so those stay out of balloons.
	 */
	beginParagraph(paragraph: Paragraph | undefined, inStory = false): void {
		this.#paragraph = paragraph && { paragraph, last: undefined, lastEnd: 0 };
		if (!paragraph || !this.active || !this.deletions || this.#deletions.has(paragraph.id)) return;
		const deletions: Deletion[] = [];
		this.#deletions.set(paragraph.id, deletions);
		if (inStory) return;
		let offset = 0;
		for (const run of paragraph.runs) {
			const length = reviewSourceLength(run);
			const revision = run.revision;
			if (revision && (revision.kind === 'delete' || revision.kind === 'moveFrom')) {
				const previous = deletions.at(-1);
				if (
					previous &&
					previous.end === offset &&
					previous.revision.author === revision.author &&
					previous.revision.kind === revision.kind
				) {
					previous.end += length;
					previous.text += run.text;
				} else deletions.push({ start: offset, end: offset + length, text: run.text, revision });
			}
			offset += length;
		}
	}

	/** One painted fragment of the current paragraph (page pixels). */
	fragment(
		run: TextRun | undefined,
		fragment: LayoutFragment,
		x: number,
		lineTop: number,
		lineHeight: number,
	): void {
		const current = this.#paragraph;
		if (!this.active || !current) return;
		const anchor = { x, top: lineTop, height: lineHeight };
		const start = fragment.sourceStart;
		if (start !== undefined) this.#flushDeletions(current.paragraph.id, start, anchor);
		if (run && this.formatting) {
			const revision =
				run.formatRevision ?? (run.revision?.kind === 'formatChange' ? run.revision : undefined);
			const key = revision && `${current.paragraph.id}:run:${revision.id}`;
			if (revision && key && !this.#seen.has(key)) {
				this.#seen.add(key);
				const items = describeRunFormatChange(run, this.model);
				if (items) this.#revision(revision, formatChangeText(items), anchor, 'format');
			}
		}
		for (const id of run?.commentIds ?? []) this.#commentAnchor(id, anchor, fragment.widthPx);
		current.last = { x: x + fragment.widthPx, top: lineTop, height: lineHeight };
		current.lastEnd = Math.max(current.lastEnd, fragment.sourceEnd ?? 0);
	}

	/** The fragment with its comment range highlighted, while balloons are shown. */
	style(fragment: LayoutFragment, run: TextRun | undefined): LayoutFragment {
		if (!this.active || fragment.background || !fragment.text) return fragment;
		const comment = this.#comments.first(run?.commentIds);
		return comment ? { ...fragment, background: paleColor(authorColor(comment.author)) } : fragment;
	}

	/** The paragraph's last box ends: its trailing deletions and its formatting change. */
	endParagraph(boxStart: Anchor): void {
		const current = this.#paragraph;
		this.#paragraph = undefined;
		if (!this.active || !current) return;
		const { paragraph } = current;
		let offset = 0;
		let visibleAfter = false;
		for (const run of paragraph.runs) {
			const length = reviewSourceLength(run);
			const hidden =
				this.deletions && (run.revision?.kind === 'delete' || run.revision?.kind === 'moveFrom');
			if (offset >= current.lastEnd && length > 0 && !hidden) visibleAfter = true;
			offset += length;
		}
		if (visibleAfter) return;
		const end = current.last ?? boxStart;
		this.#flushDeletions(paragraph.id, Infinity, end);
		const key = `${paragraph.id}:paragraph`;
		if (this.formatting && paragraph.formatRevision && !this.#seen.has(key)) {
			this.#seen.add(key);
			const items = describeParagraphFormatChange(paragraph, this.model);
			if (items) this.#revision(paragraph.formatRevision, formatChangeText(items), end, 'format');
		}
		this.section(paragraph.id, end);
	}

	/** Emits a table property change balloon (`w:tblPrChange`), returning whether one fired. */
	table(blockId: string, anchor: Anchor): boolean {
		if (!this.active || !this.formatting) return false;
		const table = this.#tables.get(blockId);
		const key = `${blockId}:table`;
		if (!table?.formatRevision || this.#seen.has(key)) return false;
		this.#seen.add(key);
		const items = describeTableFormatChange(table);
		if (!items) return false;
		this.#revision(table.formatRevision, formatChangeText(items), anchor, 'format');
		return true;
	}

	/** Emits a row property change balloon (`w:trPrChange`), returning whether one fired. */
	row(blockId: string, rowIndex: number, anchor: Anchor): boolean {
		if (!this.active || !this.formatting) return false;
		const row = this.#tables.get(blockId)?.rowProperties?.[rowIndex];
		const key = `${blockId}:row:${rowIndex}`;
		if (!row?.formatRevision || this.#seen.has(key)) return false;
		this.#seen.add(key);
		const items = describeRowFormatChange(row);
		if (!items) return false;
		this.#revision(row.formatRevision, formatChangeText(items), anchor, 'format');
		return true;
	}

	/** Emits a cell property change balloon (`w:tcPrChange`), returning whether one fired. */
	cell(blockId: string, rowIndex: number, cellIndex: number, anchor: Anchor): boolean {
		if (!this.active || !this.formatting) return false;
		const cell = this.#tables.get(blockId)?.rows[rowIndex]?.[cellIndex];
		const key = `${blockId}:cell:${rowIndex}:${cellIndex}`;
		if (!cell?.formatRevision || this.#seen.has(key)) return false;
		this.#seen.add(key);
		const items = describeCellFormatChange(cell);
		if (!items) return false;
		this.#revision(cell.formatRevision, formatChangeText(items), anchor, 'format');
		return true;
	}

	/** Emits a section property change balloon (`w:sectPrChange`) at the section's last block. */
	section(blockId: string, anchor: Anchor): boolean {
		if (!this.active || !this.formatting) return false;
		const section = this.#sections.get(blockId);
		const key = `${blockId}:section`;
		if (!section?.formatRevision || this.#seen.has(key)) return false;
		this.#seen.add(key);
		const items = describeSectionFormatChange(section);
		if (!items) return false;
		this.#revision(section.formatRevision, formatChangeText(items), anchor, 'format');
		return true;
	}

	/** Moves the comment's PDF note onto its balloon once the balloon is placed. */
	comment(id: string, note: PdfNoteAnnotation, layout: LayoutPageBox): void {
		const balloon = this.#pageComments.get(id);
		if (!balloon) return;
		const frame = balloonFrame(layout);
		// Viewers draw a note's icon about 20 points square from its top-left: inside the balloon's
		// top-right corner.
		balloon.placed = (box) => {
			const right = (box.x + box.width) * PT_PER_PX * frame.scale - 1;
			const top = frame.offsetYPt + (layout.heightPx - box.y) * PT_PER_PX * frame.scale - 1;
			note.rect = [right - 20, top - 18, right, top];
		};
	}

	/** Paints the page's pane and balloons (in the page's pixel space) and forgets them. */
	endPage(page: PdfVectorPage, layout: LayoutPageBox): void {
		const items = this.#items;
		this.#items = [];
		this.#pageComments = new Map();
		if (!this.active) return;
		if (!paintBalloons(page, layout, items, this.text))
			this.report.note(
				'Some markup balloons did not fit beside their page and overlap at the top of the markup pane.',
			);
	}

	get #label(): Omit<BalloonPiece, 'text'> {
		return { bold: true, family: this.#labelFamily };
	}

	#flushDeletions(paragraphId: string, before: number, anchor: Anchor): void {
		const pending = this.#deletions.get(paragraphId) ?? [];
		const count = pending.findIndex((deletion) => deletion.end > before);
		for (const deletion of pending.splice(0, count < 0 ? pending.length : count)) {
			const moved = deletion.revision.kind === 'moveFrom';
			this.#revision(
				deletion.revision,
				[
					{ text: moved ? 'Moved: ' : 'Deleted: ', ...this.#label },
					{ text: deletion.text, family: this.#family },
				],
				anchor,
				'delete',
				moved ? MOVE_COLOR : undefined,
			);
		}
	}

	#revision(
		revision: Revision,
		content: string | BalloonPiece[],
		anchor: Anchor,
		kind: Balloon['kind'],
		color = authorColor(revision.author),
	): void {
		const label = typeof content === 'string' ? /^[^:]*:?/.exec(content)![0] : '';
		const pieces =
			typeof content === 'string'
				? [
						{ text: label, ...this.#label },
						{ text: content.slice(label.length), family: this.#labelFamily },
					]
				: content;
		this.#items.push({
			kind,
			color,
			paragraphs: [pieces],
			anchorX: anchor.x,
			anchorTop: anchor.top,
			anchorBottom: anchor.top + anchor.height,
			marker: true,
		});
	}

	#commentAnchor(id: string, anchor: Anchor, width: number): void {
		const existing = this.#pageComments.get(id);
		if (existing) {
			// The connector leaves from the end of the range on its first line.
			if (Math.abs(existing.anchorTop - anchor.top) < 0.5)
				existing.anchorX = Math.max(existing.anchorX, anchor.x + width);
			return;
		}
		const comment = this.#comments.thread(id);
		if (!comment || this.#seen.has(`comment:${id}`)) return;
		this.#seen.add(`comment:${id}`);
		const color = authorColor(comment.author);
		const balloon: Balloon = {
			kind: 'comment',
			color,
			fill: paleColor(color),
			paragraphs: this.#comments.paragraphs(comment),
			anchorX: anchor.x + width,
			anchorTop: anchor.top,
			anchorBottom: anchor.top + anchor.height,
			marker: false,
		};
		this.#pageComments.set(id, balloon);
		this.#items.push(balloon);
	}
}
