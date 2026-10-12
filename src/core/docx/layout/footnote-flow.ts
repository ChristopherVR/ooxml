import type { LayoutFootnoteBox, LayoutPageBox, LayoutParagraphBox } from './result';

function firstLineBottom(note: LayoutFootnoteBox | undefined): number {
	const first = note?.paragraphs[0];
	return first
		? first.yPx + (first.lines[0]?.yPx ?? 0) + (first.lines[0]?.heightPx ?? first.heightPx)
		: (note?.heightPx ?? 0);
}

/** Space for the short rule above a page's footnotes. */
export const FOOTNOTE_SEPARATOR_PX = 12;

/** Split a note at a complete line boundary, retaining its identity and source ranges. */
function splitNote(
	note: LayoutFootnoteBox,
	budget: number,
): {
	before?: LayoutFootnoteBox;
	after?: LayoutFootnoteBox;
} {
	if (note.heightPx <= budget) return { before: note };
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let origin: number | undefined;
	for (const paragraph of note.paragraphs) {
		if (origin !== undefined) {
			after.push({ ...paragraph, yPx: paragraph.yPx - origin });
			continue;
		}
		if (paragraph.yPx + paragraph.heightPx <= budget) {
			before.push(paragraph);
			continue;
		}
		const count = paragraph.lines.filter(
			(line) => paragraph.yPx + line.yPx + line.heightPx <= budget,
		).length;
		const head = paragraph.lines.slice(0, count);
		const tail = paragraph.lines.slice(count);
		if (head.length) {
			const bottom = head.at(-1)!.yPx + head.at(-1)!.heightPx;
			before.push({ ...paragraph, heightPx: bottom, lines: head });
		}
		const first = tail[0];
		origin = paragraph.yPx + (first?.yPx ?? 0);
		after.push({
			...paragraph,
			yPx: 0,
			heightPx: Math.max(0, paragraph.heightPx - (first?.yPx ?? 0)),
			lines: tail.map((line) => ({ ...line, yPx: line.yPx - (first?.yPx ?? 0) })),
		});
	}
	if (!before.length) return { after: note };
	const last = before.at(-1)!;
	return {
		before: { ...note, heightPx: last.yPx + last.heightPx, paragraphs: before },
		...(after.length
			? {
					after: {
						...note,
						yPx: 0,
						heightPx: Math.max(0, note.heightPx - (origin ?? 0)),
						paragraphs: after,
						continued: true,
					},
				}
			: {}),
	};
}

/** Page-wide note reservation and continuation. The source document is never changed. */
export class FootnoteFlow {
	private pending:
		| { blockId: string; notes: LayoutFootnoteBox[]; bodyMinimum: number; freshBodyMinimum: number }
		| undefined;
	private remaining: LayoutFootnoteBox[] = [];

	constructor(
		private readonly report: (message: string) => void,
		private readonly separator?: (
			page: LayoutPageBox,
			continued: boolean,
		) => LayoutPageBox['footnoteSeparator'],
	) {}

	private separatorHeight(page: LayoutPageBox, notes = this.pending?.notes ?? []): number {
		return (
			(page.footnoteSeparator ?? this.separator?.(page, !!notes[0]?.continued))?.heightPx ??
			FOOTNOTE_SEPARATOR_PX
		);
	}

	get hasContinuation(): boolean {
		return this.remaining.length > 0;
	}

	hold(
		blockId: string,
		notes: LayoutFootnoteBox[],
		bodyMinimum: number,
		freshBodyMinimum = bodyMinimum,
	): void {
		this.pending = notes.length ? { blockId, notes, bodyMinimum, freshBodyMinimum } : undefined;
	}

	reserved(page: LayoutPageBox): number {
		return page.footnotes?.length
			? this.separatorHeight(page) + page.footnotes.reduce((sum, note) => sum + note.heightPx, 0)
			: 0;
	}

	pendingHeight(page: LayoutPageBox, capacity: number, bodyUsed: number): number {
		if (!this.pending) return 0;
		const total = this.pending.notes.reduce((sum, note) => sum + note.heightPx, 0);
		const separator = page.footnotes?.length ? 0 : this.separatorHeight(page);
		// Short notes retain the existing keep-with-reference behavior. A note larger than a fresh
		// page must share that page with at least the reference's first line, then continue.
		if (total + this.separatorHeight(page) + this.pending.freshBodyMinimum <= capacity)
			return total + separator;
		return Math.max(
			firstLineBottom(this.pending.notes[0]) + separator,
			capacity - bodyUsed - this.reserved(page) - this.pending.bodyMinimum,
		);
	}

	/** A later column cannot grow the page-wide note area into body content already placed. */
	blockedByEarlierColumn(page: LayoutPageBox, capacity: number, bodyBottom: number): boolean {
		if (!this.pending) return false;
		const total = this.pending.notes.reduce((sum, note) => sum + note.heightPx, 0);
		const minimum = firstLineBottom(this.pending.notes[0]);
		const needed =
			total + this.separatorHeight(page) + this.pending.freshBodyMinimum <= capacity
				? total
				: minimum;
		return (
			bodyBottom +
				this.reserved(page) +
				needed +
				(page.footnotes?.length ? 0 : this.separatorHeight(page)) >
			capacity
		);
	}

	commit(blockId: string, page: LayoutPageBox, available: number): void {
		if (this.pending?.blockId !== blockId) return;
		const notes = this.pending.notes;
		this.pending = undefined;
		this.append(page, notes, available);
	}

	resume(page: LayoutPageBox, capacity: number): void {
		const notes = this.remaining;
		this.remaining = [];
		this.append(page, notes, capacity);
	}

	private append(page: LayoutPageBox, notes: LayoutFootnoteBox[], available: number): void {
		if (!page.footnotes?.length && notes.length) {
			const separator = this.separator?.(page, !!notes[0]?.continued);
			if (separator) page.footnoteSeparator = separator;
		}
		let top = page.footnotes?.reduce((sum, note) => sum + note.heightPx, 0) ?? 0;
		let budget = Math.max(0, available - (page.footnotes?.length ? 0 : this.separatorHeight(page)));
		for (let index = 0; index < notes.length; index++) {
			const note = notes[index]!;
			let split = splitNote(note, budget);
			if (!split.before && !page.footnotes?.length) {
				split = splitNote(note, firstLineBottom(note));
				this.report(
					'A footnote line was taller than its available page area and was placed without splitting.',
				);
			}
			if (split.before) {
				(page.footnotes ??= []).push({ ...split.before, yPx: top });
				top += split.before.heightPx;
				budget -= split.before.heightPx;
			}
			if (split.after) {
				this.remaining.push(split.after, ...notes.slice(index + 1));
				break;
			}
		}
	}
}
