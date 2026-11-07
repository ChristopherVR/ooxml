/**
 * table-cell-paragraphs.ts - lay a table cell's text out paragraph by paragraph.
 *
 * Lays out a cell with `PptxTableCell.paragraphs` as one block per paragraph,
 * using the CSS shape text uses. Bindings map these blocks onto their own
 * markup.
 *
 * @module render/table-cell-paragraphs
 */
import type { PptxTableCell, PptxTableCellParagraph } from 'ooxml-core/pptx';

import { resolveParagraphIndent } from './bullet-list';
import type { FontBoxMeasure } from './paragraph-baseline';
import { armBaselineSettle } from './paragraph-baseline-settle';
import { resolveParagraphSpacing } from './paragraph-spacing';
import { cellBaselinePlacement } from './table-cell-baseline';
import type { CellTextRun, TableCellCss } from './table-style';
import { DEFAULT_LINE_HEIGHT } from './text-line-height';
import {
	resolveCssTextAlign,
	resolveTextAlignLast,
	DISTRIBUTED_TEXT_JUSTIFY,
} from './text-paragraph-style';

/** One paragraph of a cell: its block CSS and the runs inside it. */
export interface CellParagraphBlock {
	css: TableCellCss;
	runs: CellTextRun[];
}

/** A paragraph's layout CSS, and the `line-height` its runs take, if any. */
interface ParagraphLayout {
	css: TableCellCss;
	runLineHeight?: number;
	/** The paragraph's own space before (px), which the PREVIOUS block's margin carries. */
	spaceBeforePx?: number;
}

/** An empty paragraph still takes a line, as it does in PowerPoint. */
const BLANK_LINE: CellTextRun = { text: '', isLineBreak: true };

/** What a caller can tell {@link cellParagraphBlocks} about the cell it paints. */
export interface CellBlockOptions {
	cellCss?: TableCellCss | undefined;
	defaultFontPt?: number | undefined;
	measure?: FontBoxMeasure | undefined;
}

/** Whether a paragraph's runs hold any text (an empty paragraph keeps the browser's blank line). */
function hasContentRuns(runs: readonly CellTextRun[]): boolean {
	return runs.some((run) => run.text !== '' && !run.isLineBreak);
}

/** Split runs at their paragraph markers. */
function groupRuns(runs: readonly CellTextRun[]): CellTextRun[][] {
	const groups: CellTextRun[][] = [[]];
	for (const run of runs) {
		if (run.isParagraphBreak) {
			groups.push([]);
		} else {
			groups[groups.length - 1].push(run);
		}
	}
	return groups;
}

/**
 * The smallest run size, like `resolveParagraphStrutFontSize` for shape text,
 * so the paragraph's line box rests on its own text rather than the cell's.
 * Only when every run with text sets a size: a cell run without one inherits
 * the block's.
 */
function strutFontSizePt(runs: readonly CellTextRun[]): number | undefined {
	let smallest: number | undefined;
	for (const run of runs) {
		if (!run.text) {
			continue;
		}
		if (typeof run.fontSize !== 'number' || run.fontSize <= 0) {
			return undefined;
		}
		smallest = smallest === undefined ? run.fontSize : Math.min(smallest, run.fontSize);
	}
	return smallest;
}

/**
 * The layout CSS one paragraph sets, before any line box sizing. Margins are
 * logical so they follow a vertical or right-to-left cell. A paragraph that
 * follows the cell's alignment sets none of its own.
 */
function paragraphLayoutCss(
	paragraph: PptxTableCellParagraph | undefined,
	followsCell: boolean,
	isFirst: boolean,
	isLast: boolean,
): ParagraphLayout {
	const css: TableCellCss = {};
	const rtl = paragraph?.rtl === true;
	if (!followsCell) {
		// A paragraph that sets no alignment must not take the cell's, which
		// comes from its first paragraph.
		css.textAlign = resolveCssTextAlign(paragraph?.align, rtl) ?? 'start';
		const textAlignLast = resolveTextAlignLast(paragraph?.align);
		if (textAlignLast !== undefined) {
			css.textAlignLast = textAlignLast;
			css.textJustify = DISTRIBUTED_TEXT_JUSTIFY;
		}
	}
	if (paragraph?.rtl !== undefined) {
		// As shape text does: the paragraph sets its own BiDi embedding level.
		css.direction = rtl ? 'rtl' : 'ltr';
		css.unicodeBidi = 'embed';
	}

	const indent = resolveParagraphIndent(
		{ marginLeft: paragraph?.paragraphMarginLeft, indent: paragraph?.paragraphIndent },
		undefined,
	);
	if (indent.marginLeftPx !== undefined) {
		css.marginInlineStart = `${indent.marginLeftPx}px`;
	}
	// A negative indent hangs the first line into the margin; with no margin to
	// hang into it would push the text out of the cell.
	if (
		indent.textIndentPx !== undefined &&
		(indent.textIndentPx > 0 || indent.marginLeftPx !== undefined)
	) {
		css.textIndent = `${indent.textIndentPx}px`;
	}

	const spacing = resolveParagraphSpacing({ paraProps: paragraph, isFirst, isLast });
	if (spacing.lineHeight !== undefined) {
		css.lineHeight = spacing.lineHeight;
	}
	if (spacing.spaceAfterPx !== undefined) {
		css.marginBlockEnd = `${spacing.spaceAfterPx}px`;
	}
	const layout: ParagraphLayout = { css };
	if (spacing.runLineHeight !== undefined) {
		layout.runLineHeight = spacing.runLineHeight;
	}
	if (spacing.spaceBeforePx !== undefined) {
		layout.spaceBeforePx = spacing.spaceBeforePx;
	}
	return layout;
}

/**
 * Move each paragraph's own space before onto the previous block's
 * `margin-block-end`. PowerPoint adds a paragraph's `spcAft` and the next
 * one's `spcBef`; the cell stacks its blocks in normal flow, where two
 * adjacent margins would collapse to the larger, so the sum goes on one side.
 */
function addSpaceBefore(layouts: ParagraphLayout[]): void {
	for (let index = 1; index < layouts.length; index++) {
		const before = layouts[index].spaceBeforePx;
		if (before === undefined) {
			continue;
		}
		const previous = layouts[index - 1].css;
		const after = Number.parseFloat(String(previous.marginBlockEnd ?? '0')) || 0;
		previous.marginBlockEnd = `${after + before}px`;
	}
}

/** A paragraph's alignment, with PowerPoint's default. */
function alignOf(paragraph: PptxTableCellParagraph | undefined): string {
	return paragraph?.align ?? 'left';
}

/**
 * Whether a paragraph takes the cell's own alignment: it aligns like the first
 * paragraph, which the cell's alignment comes from. Its block then sets none,
 * so it follows the cell (`style.align`, `anchorCtr`), including after the
 * cell's alignment is edited. A right-to-left paragraph follows only an
 * explicit `style.align`; otherwise it aligns right by its direction.
 */
function followsCellAlign(
	cell: PptxTableCell,
	paragraph: PptxTableCellParagraph | undefined,
	first: PptxTableCellParagraph | undefined,
): boolean {
	return (
		alignOf(paragraph) === alignOf(first) &&
		(paragraph?.rtl !== true || cell.style?.align !== undefined)
	);
}

/**
 * Whether the cell needs per-paragraph blocks: some paragraph is right to
 * left, sets an indent, spacing or a line spacing other than single, or an
 * alignment other than the first paragraph's. Writers that repeat the
 * defaults on every paragraph (Google Slides writes `algn="l" rtl="0"`,
 * LibreOffice single line spacing) keep the run stream.
 */
function needsBlocks(
	paragraphs: readonly (PptxTableCellParagraph | undefined)[],
	follows: readonly boolean[],
	layouts: readonly ParagraphLayout[],
	placed: readonly boolean[],
): boolean {
	return follows.some(
		(followsCell, index) =>
			!followsCell ||
			placed[index] === true ||
			paragraphs[index]?.rtl === true ||
			layouts[index].css.marginInlineStart !== undefined ||
			layouts[index].css.textIndent !== undefined ||
			layouts[index].css.marginBlockEnd !== undefined ||
			(layouts[index].css.lineHeight !== undefined &&
				layouts[index].css.lineHeight !== DEFAULT_LINE_HEIGHT),
	);
}

/**
 * Resolve a cell's paragraphs into blocks.
 *
 * A cell edited to plain text has no runs; its text is then one run in its
 * one remaining paragraph.
 *
 * @param cell - The cell being painted.
 * @param options - `cellCss`, the cell's resolved style (its font and line
 *   height, which a paragraph without its own inherits), and `defaultFontPt`,
 *   the table's default text size: with them a paragraph is moved onto
 *   PowerPoint's baselines (see `table-cell-baseline`). `measure` replaces the
 *   canvas font measurement in tests.
 * @returns One block per paragraph, or `undefined` when the cell has no
 *   paragraph layout or nothing in it needs blocks, in which case the binding
 *   draws its runs as before.
 */
export function cellParagraphBlocks(
	cell: PptxTableCell,
	options: CellBlockOptions = {},
): CellParagraphBlock[] | undefined {
	const { paragraphs } = cell;
	if (!paragraphs || paragraphs.length === 0) {
		return undefined;
	}
	const runs: CellTextRun[] =
		cell.textRuns && cell.textRuns.length > 0 ? cell.textRuns : [{ text: cell.text }];
	const groups = groupRuns(runs);
	const follows = groups.map((_, index) =>
		followsCellAlign(cell, paragraphs[index], paragraphs[0]),
	);
	const layouts = groups.map((_, index) =>
		paragraphLayoutCss(paragraphs[index], follows[index], index === 0, index === groups.length - 1),
	);
	addSpaceBefore(layouts);
	// Exact line spacing: the runs must not grow the line, as in shape text.
	const laidOut = groups.map((group, index) => {
		const { runLineHeight } = layouts[index];
		return runLineHeight === undefined
			? group
			: group.map((run) => ({ ...run, lineHeight: runLineHeight }));
	});
	// Placed from the runs as they are laid out, line heights included.
	const placements = groups.map((group, index) =>
		hasContentRuns(group)
			? cellBaselinePlacement({
					runs: laidOut[index],
					lineHeight: layouts[index].css.lineHeight,
					strutPt: strutFontSizePt(group),
					cellCss: options.cellCss,
					defaultFontPt: options.defaultFontPt,
					measure: options.measure,
				})
			: undefined,
	);
	if (
		!needsBlocks(
			paragraphs,
			follows,
			layouts,
			placements.map((placement) => placement !== undefined),
		)
	) {
		return undefined;
	}
	return groups.map((group, index) => {
		const { css } = layouts[index];
		const placement = placements[index];
		if (placement) {
			Object.assign(css, placement.style);
			if (placement.extentDelta !== undefined && index === groups.length - 1) {
				// The cell ends at the last line's baseline plus descent, not at its pitch.
				const after = Number.parseFloat(String(css.marginBlockEnd ?? '0')) || 0;
				css.marginBlockEnd = `${Math.round((after + placement.extentDelta) * 100) / 100}px`;
			}
			if (placement.provisional) {
				armBaselineSettle();
			}
		}
		const hasContent = group.some((run) => run.text !== '' || run.isLineBreak);
		if (!hasContent) {
			const endPt = paragraphs[index]?.endParaFontSize;
			if (endPt !== undefined) {
				css.fontSize = `${endPt}pt`;
			}
			return { css, runs: [BLANK_LINE] };
		}
		const strutPt = strutFontSizePt(group);
		if (strutPt !== undefined) {
			css.fontSize = `${strutPt}pt`;
		}
		return { css, runs: laidOut[index] };
	});
}
