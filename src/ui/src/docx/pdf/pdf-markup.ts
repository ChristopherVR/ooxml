/**
 * "Document showing markup" for the Word PDF export, drawn the way Print Layout's All Markup
 * view draws it: inserted text underlined and deleted text struck through in its author's
 * colour (moved text in green), a change bar in the left margin beside each changed line
 * (formatting changes included), and comments as PDF notes at the start of the text they cover.
 * Balloons in the markup pane are `pdf-balloon-items.ts`. DOM-free.
 */
import type { Comment, Paragraph, TextRun } from 'ooxml-core/docx';
import { authorColor } from 'ooxml-core/docx/ui';
import type { LayoutFragment } from 'ooxml-core/docx/layout';
import type { PdfNoteAnnotation, PdfVectorPage } from '../../export/pdf-vector';

const MOVE_COLOR = '#2e7d32';

/** `fragment` styled for its run's tracked change, or as it is when the run has none. */
export function markupFragment(fragment: LayoutFragment, run: TextRun | undefined): LayoutFragment {
	const kind = run?.revision?.kind;
	if (
		!run?.revision ||
		(kind !== 'insert' && kind !== 'delete' && kind !== 'moveFrom' && kind !== 'moveTo')
	)
		return fragment;
	const color =
		kind === 'moveFrom' || kind === 'moveTo' ? MOVE_COLOR : authorColor(run.revision.author);
	const inserted = kind === 'insert' || kind === 'moveTo';
	return {
		...fragment,
		color,
		...(inserted ? { underline: true } : { strike: true }),
	};
}

/** Whether a run is part of a tracked insertion, deletion or move (it gets a change bar). */
export const isChanged = (run: TextRun | undefined): boolean => {
	const kind = run?.revision?.kind;
	return kind === 'insert' || kind === 'delete' || kind === 'moveFrom' || kind === 'moveTo';
};

/**
 * Whether a run's formatting, or its paragraph's, is a tracked change (`w:rPrChange`,
 * `w:pPrChange`): Word gives those lines a change bar too.
 */
export const isFormatChanged = (
	run: TextRun | undefined,
	paragraph: Paragraph | undefined,
): boolean =>
	run?.formatRevision?.kind === 'formatChange' ||
	run?.revision?.kind === 'formatChange' ||
	paragraph?.formatRevision?.kind === 'paragraphChange';

/** The change bars of one page: one per changed line, in the left margin. */
export class DocxPdfChangeBars {
	readonly #lines = new Map<string, { top: number; height: number }>();
	get size(): number {
		return this.#lines.size;
	}
	add(top: number, height: number): void {
		this.#lines.set(`${top.toFixed(2)}:${height.toFixed(2)}`, { top, height });
	}
	/** Paints the bars (y-down CSS pixels) and forgets them. */
	paint(page: PdfVectorPage, marginLeftPx: number): void {
		if (!this.#lines.size) return;
		const x = Math.max(2, marginLeftPx - 24);
		page.fillColor(0, 0, 0);
		for (const { top, height } of this.#lines.values()) page.rect(x, top, 1, height);
		page.op('f');
		this.#lines.clear();
	}
}

/** The note a comment becomes, its icon at a point given in PDF points (y up). */
export function commentNote(comment: Comment, x: number, top: number): PdfNoteAnnotation {
	return {
		kind: 'note',
		rect: [x, top - 18, x + 20, top],
		contents: comment.text,
		...(comment.author ? { author: comment.author } : {}),
		...(comment.date ? { date: comment.date } : {}),
	};
}
