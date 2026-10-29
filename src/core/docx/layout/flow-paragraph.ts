import type { PageCursor } from './page-cursor';
import { adjustForWidowOrphan, widowControlEnabled } from './keep-rules';
import type { ParagraphLayoutResult } from './paragraph-layout';
import type { LayoutParagraph } from './input';
import type { LayoutFootnoteBox, LayoutLine, LayoutParagraphBox } from './result';
import { expectDefined } from '../expect-defined';

export interface ParagraphPlacement {
	/** Measured Word compatibility modes do not reserve a text line for an empty column-break marker. */
	trimLeadingColumnBreak?: boolean;
	paragraph: LayoutParagraph;
	layout: ParagraphLayoutResult;
	/** Spacing-before in px, with `w:contextualSpacing` suppression already resolved by the caller. */
	spacingBeforePx: number;
	/** Compatibility modes before 15 keep the space before of a paragraph that starts a page by `pageBreakBefore`. */
	keepSpaceBeforePageBreak?: boolean;
	/** Spacing-after in px when it differs from the paragraph's own (contextual spacing). */
	spacingAfterPx?: number;
	/**
	 * Minimum height (including this paragraph and any `keepNext` chain after
	 * it, as resolved by the caller) that must fit before this paragraph may
	 * start; 0 when neither `keepNext` nor `keepLines` apply.
	 */
	requiredTogetherPx: number;
	/** Lays unconsumed tokens out at the cursor's new column width. */
	reflow?: (startToken: number) => ParagraphLayoutResult;
	/** Notes referenced by these lines, laid out at the cursor's current column width. */
	footnotes?: (lines: readonly LayoutLine[], includeUnanchored: boolean) => LayoutFootnoteBox[];
}

/**
 * Places one paragraph's lines into `cursor`, splitting across page/column
 * boundaries as needed. Honors `pageBreakBefore`, the keep-together height
 * computed by the caller (chained `keepNext`/whole-paragraph `keepLines`),
 * default 2-line widow/orphan control, and any explicit page/column break
 * tokens carried on the paragraph's runs.
 */
export function placeParagraph(
	cursor: PageCursor,
	placement: ParagraphPlacement,
	note: (message: string) => void,
): void {
	const { paragraph, requiredTogetherPx, reflow } = placement;
	const spacingAfterPx = placement.spacingAfterPx ?? placement.layout.spacingAfterPx;
	// Compatibility modes before 15 keep the space before of a `pageBreakBefore` paragraph.
	const dropsSpaceBefore = () =>
		cursor.atContinuationTop && !(paragraph.pageBreakBefore && placement.keepSpaceBeforePageBreak);
	let spacingBeforePx = dropsSpaceBefore() ? 0 : placement.spacingBeforePx;
	// Native modes 12/14/15 retain the empty paragraph line after a terminal column break.
	const trailingBreakLine = (result: ParagraphLayoutResult): ParagraphLayoutResult => {
		const last = result.lines.at(-1);
		if (
			!placement.trimLeadingColumnBreak ||
			!last ||
			!result.columnBreakAfterLine.has(result.lines.length - 1)
		)
			return result;
		return {
			...result,
			contentHeightPx: result.contentHeightPx + last.heightPx,
			lines: [
				...result.lines,
				{
					...last,
					yPx: last.yPx + last.heightPx,
					fragments: [],
					sourceStart: last.sourceEnd,
					sourceEnd: last.sourceEnd,
				},
			],
		};
	};
	let layout = trailingBreakLine(placement.layout);
	let width = cursor.columnWidthPx;
	if (paragraph.pageBreakBefore && cursor.hasPageContent) cursor.newPage();
	if (requiredTogetherPx > 0 && placement.footnotes) {
		cursor.holdFootnotes(
			paragraph.id,
			placement.footnotes(layout.lines, true),
			requiredTogetherPx,
			layout.lines[0]?.heightPx ?? 0,
		);
	}
	if (
		requiredTogetherPx > 0 &&
		requiredTogetherPx <= cursor.columnHeightPx &&
		requiredTogetherPx > cursor.remainingHeightPx() &&
		!cursor.atColumnTop
	) {
		cursor.newColumn();
	}
	if (dropsSpaceBefore()) spacingBeforePx = 0;

	const widow = widowControlEnabled(paragraph.widowControl);
	let lineCursor = 0;
	let nextToken = 0;
	let firstSegment = true;
	const lineHeight = (line: LayoutLine, index: number) =>
		placement.trimLeadingColumnBreak &&
		layout.columnBreakAfterLine.has(index) &&
		line.fragments.every((fragment) => !fragment.text && !fragment.object)
			? 0
			: line.heightPx;
	const placedNotes = new Set<string>();
	const notesFor = (lines: readonly LayoutLine[]): LayoutFootnoteBox[] =>
		(placement.footnotes?.(lines, firstSegment) ?? []).filter((item) => !placedNotes.has(item.id));
	const hold = (lines: readonly LayoutLine[], bodyMinimum: number): LayoutFootnoteBox[] => {
		const notes = notesFor(lines);
		if (placement.footnotes)
			cursor.holdFootnotes(
				paragraph.id,
				notes,
				bodyMinimum,
				Math.max(0, ...lines.map((line) => line.heightPx)),
			);
		return notes;
	};
	const reflowIfNeeded = () => {
		if (!reflow || cursor.columnWidthPx === width) return;
		layout = trailingBreakLine(reflow(nextToken));
		lineCursor = 0;
		width = cursor.columnWidthPx;
	};
	reflowIfNeeded();
	if (!layout.lines.length) {
		// An empty paragraph still occupies its spacing/blank-line height.
		cursor.place(
			{ kind: 'paragraph', blockId: paragraph.id, yPx: 0, heightPx: 0, lines: [] },
			spacingBeforePx + spacingAfterPx,
			spacingBeforePx,
		);
		return;
	}
	while (lineCursor < layout.lines.length) {
		reflowIfNeeded();
		hold(
			layout.lines.slice(lineCursor, lineCursor + 1),
			layout.lines[lineCursor] ? lineHeight(layout.lines[lineCursor]!, lineCursor) : 0,
		);
		// A continuation can fill an otherwise empty physical page. Advance before applying the
		// oversized-body-line fallback; the notes have made progress even though the body has not.
		cursor.advancePastFootnotes(
			layout.lines[lineCursor] ? lineHeight(layout.lines[lineCursor]!, lineCursor) : 0,
		);
		reflowIfNeeded();
		let cumulative = 0;
		let end = lineCursor;
		let forced: 'page' | 'column' | undefined;
		for (let i = lineCursor; i < layout.lines.length; i++) {
			// A line's gap (clearing a wrapped picture) moves with it, except when a later segment
			// starts it at the top of a new column or page.
			const line = expectDefined(layout.lines[i], 'paragraph line index');
			const gap = i === lineCursor && !firstSegment ? 0 : (line.gapBeforePx ?? 0);
			const height = lineHeight(line, i) + gap;
			hold(
				layout.lines.slice(lineCursor, i + 1),
				cumulative + height + (firstSegment ? spacingBeforePx : 0),
			);
			const budget = cursor.remainingHeightPx() - (firstSegment ? spacingBeforePx : 0);
			// A line that does not fit waits for the next column, unless the column is empty.
			if (cumulative + height > budget + 1e-6 && (end > lineCursor || !cursor.atColumnTop)) break;
			cumulative += height;
			end = i + 1;
			if (layout.pageBreakAfterLine.has(i)) {
				forced = 'page';
				break;
			}
			if (layout.columnBreakAfterLine.has(i)) {
				forced = 'column';
				break;
			}
		}
		let placedCount = end - lineCursor;
		if (!forced && end < layout.lines.length)
			placedCount = adjustForWidowOrphan(layout.lines.length - lineCursor, placedCount, widow);
		if (placedCount === 0) {
			if (!cursor.atColumnTop) {
				cursor.newColumn();
				if (firstSegment && dropsSpaceBefore()) spacingBeforePx = 0;
				continue;
			}
			note(
				'A paragraph line was taller than an empty page/column and was placed without splitting.',
			);
			placedCount = Math.max(1, end - lineCursor);
		}
		const committedNotes = hold(
			layout.lines.slice(lineCursor, lineCursor + placedCount),
			cumulative + (firstSegment ? spacingBeforePx : 0),
		);
		const first = expectDefined(layout.lines[lineCursor], 'first line of paragraph segment');
		const origin = first.yPx - (firstSegment ? (first.gapBeforePx ?? 0) + layout.insetTopPx : 0);
		const segmentLines = layout.lines
			.slice(lineCursor, lineCursor + placedCount)
			.map((line, index) => ({
				...line,
				yPx: line.yPx - origin,
				heightPx: lineHeight(line, lineCursor + index),
			}));
		const isFinal = lineCursor + placedCount >= layout.lines.length;
		const lastLine = segmentLines.at(-1);
		const isLastSegment = lineCursor + placedCount >= layout.lines.length;

		const segmentHeight =
			(lastLine ? lastLine.yPx + lastLine.heightPx : 0) +
			(isLastSegment ? layout.insetBottomPx : 0);
		const box: LayoutParagraphBox = {
			kind: 'paragraph',
			blockId: paragraph.id,
			yPx: 0,
			heightPx: segmentHeight,
			lines: segmentLines,
			...(layout.frame ? { frame: layout.frame } : {}),
		};
		const advance =
			(firstSegment ? spacingBeforePx : 0) + segmentHeight + (isFinal ? spacingAfterPx : 0);
		cursor.place(box, advance, firstSegment ? spacingBeforePx : 0);
		for (const item of committedNotes) placedNotes.add(item.id);
		nextToken = layout.lines[lineCursor + placedCount - 1]?.nextToken ?? nextToken;
		lineCursor += placedCount;
		firstSegment = false;
		if (forced === 'page') cursor.newPage();
		else if (forced === 'column') cursor.newColumn();
		else if (lineCursor < layout.lines.length) cursor.newColumn();
	}
}
