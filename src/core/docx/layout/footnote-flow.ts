import type { LayoutFootnoteBox, LayoutPageBox } from './result';
import type { LayoutColumns } from './input';
import { splitNote } from './footnote-split';
import {
	balancedFootnotes,
	footnoteColumns,
	minimumFootnoteHeight,
	type ReflowFootnote,
} from './footnote-columns';

function firstLineBottom(note: LayoutFootnoteBox | undefined): number {
	const first = note?.paragraphs[0];
	return first
		? first.yPx + (first.lines[0]?.yPx ?? 0) + (first.lines[0]?.heightPx ?? first.heightPx)
		: (note?.heightPx ?? 0);
}

/** Space for the short rule above a page's footnotes. */
export const FOOTNOTE_SEPARATOR_PX = 12;

/** Page-wide note reservation and continuation. The source document is never changed. */
export class FootnoteFlow {
	private pending:
		| { blockId: string; notes: LayoutFootnoteBox[]; bodyMinimum: number; freshBodyMinimum: number }
		| undefined;
	private remaining: LayoutFootnoteBox[] = [];
	private sources = new WeakMap<LayoutPageBox, LayoutFootnoteBox[]>();

	constructor(
		private readonly report: (message: string) => void,
		private readonly separator?: (
			page: LayoutPageBox,
			continued: boolean,
		) => LayoutPageBox['footnoteSeparator'],
		private readonly columns?: LayoutColumns,
		private readonly reflow: ReflowFootnote = (note) => note,
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

	reserved(page: LayoutPageBox, _columnIndex = 0): number {
		return page.footnotes?.length
			? this.separatorHeight(page) +
					(page.footnoteHeightPx ?? page.footnotes.reduce((sum, note) => sum + note.heightPx, 0))
			: 0;
	}

	pendingHeight(page: LayoutPageBox, capacity: number, bodyUsed: number, _columnIndex = 0): number {
		if (!this.pending) return 0;
		if (this.columns) {
			const existing = this.sources.get(page) ?? [];
			const height =
				minimumFootnoteHeight(
					[...existing, ...this.pending.notes],
					footnoteColumns(page, this.columns),
					this.reflow,
				) + this.separatorHeight(page);
			if (height + this.pending.freshBodyMinimum <= capacity)
				return Math.max(0, height - this.reserved(page));
			return Math.max(
				firstLineBottom(this.pending.notes[0]) +
					(page.footnotes?.length ? 0 : this.separatorHeight(page)),
				capacity - bodyUsed - this.reserved(page) - this.pending.bodyMinimum,
			);
		}
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
		if (this.columns)
			return (
				bodyBottom + this.reserved(page) + this.pendingHeight(page, capacity, bodyBottom) > capacity
			);
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

	commit(blockId: string, page: LayoutPageBox, available: number, _columnIndex = 0): boolean {
		if (this.pending?.blockId !== blockId) return false;
		const notes = this.pending.notes;
		this.pending = undefined;
		this.append(page, notes, available);
		return true;
	}

	resume(page: LayoutPageBox, capacity: number, _columnIndex?: number, width?: number): void {
		const notes =
			width === undefined ? this.remaining : this.remaining.map((note) => this.reflow(note, width));
		this.remaining = [];
		this.append(page, notes, capacity);
	}

	private append(page: LayoutPageBox, notes: LayoutFootnoteBox[], available: number): void {
		if (!page.footnotes?.length && notes.length) {
			const separator = this.separator?.(page, !!notes[0]?.continued);
			if (separator) page.footnoteSeparator = separator;
		}
		if (this.columns) {
			const sources = [...(this.sources.get(page) ?? []), ...notes];
			this.sources.set(page, sources);
			const geometry = footnoteColumns(page, this.columns);
			const previousHeight = page.footnoteHeightPx ?? 0;
			const capacity = Math.max(
				0,
				available + previousHeight - (page.footnotes?.length ? 0 : this.separatorHeight(page)),
			);
			let packed = balancedFootnotes(sources, geometry, capacity, this.reflow);
			if (!packed.placed.length && sources.length) {
				this.report(
					'A footnote line was taller than its available page area and was placed without splitting.',
				);
				packed = balancedFootnotes(
					sources,
					geometry,
					firstLineBottom(this.reflow(sources[0]!, geometry[0]!.widthPx)),
					this.reflow,
				);
			}
			page.footnotes = packed.placed;
			page.footnoteHeightPx = packed.heightPx;
			this.remaining = packed.remaining.map((note) => ({
				...note,
				continued: !!note.continued || packed.placed.some((fragment) => fragment.id === note.id),
			}));
			return;
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
