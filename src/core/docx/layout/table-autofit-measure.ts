// Content widths an AutoFit table sizes its columns from (ECMA-376 17.4.53 and 17.18.87): for each
// cell, the minimum (its widest unbreakable piece of text, so nothing has to break inside a word)
// and the maximum (its text laid out without wrapping), both including paragraph indents and the
// cell margins.
import type { LayoutBlock, LayoutParagraph, LayoutTableCell } from './input';
import type { TextMeasurer } from './measure';
import { fontOf, resolveIndents, runTextWidth, tokenizeParagraph } from './paragraph-tokens';
import { isEastAsianChar, type BreakToken } from './text-breaks';
import { placeTab } from './tab-stops';

export interface ContentRange {
	minPx: number;
	maxPx: number;
}

/** One twip in CSS pixels; Word sizes columns in whole twips, rounding content widths up. */
export const TWIP_PX = 1 / 15;
/** Content width of a cell holding only an empty paragraph (measured: 6 twips). */
const EMPTY_CONTENT_PX = 6 * TWIP_PX;

const ceilTwips = (px: number) => Math.ceil(px / TWIP_PX - 1e-6) * TWIP_PX;

/** Whether text may break between two adjacent word tokens (no space between them). */
function breaksBetween(previous: BreakToken, next: BreakToken): boolean {
	if (previous.kind !== 'word' || next.kind !== 'word') return true;
	// The tokenizer only splits a run's word after a hyphen or around East Asian characters.
	if (previous.runIndex === next.runIndex) return true;
	const last = [...previous.text].at(-1) ?? '';
	const firstChar = [...next.text][0] ?? '';
	return last === '-' || isEastAsianChar(last) || isEastAsianChar(firstChar);
}

/** A paragraph's minimum and maximum widths, including its indents. */
export function paragraphRange(paragraph: LayoutParagraph, measurer: TextMeasurer): ContentRange {
	const { leftPx, rightPx, firstLineExtraPx } = resolveIndents(paragraph);
	const fonts = paragraph.runs.map(fontOf);
	const { tokens } = tokenizeParagraph(paragraph);
	const widthOf = (token: BreakToken): number => {
		const run = paragraph.runs[token.runIndex];
		if (!run) return 0;
		if (token.kind === 'object') return run.object?.widthPx ?? 0;
		if (token.kind !== 'word' && token.kind !== 'space') return 0;
		if (run.marker?.widthPx !== undefined && token.kind === 'word' && token.sourceStart === 0)
			if (token.text === run.text.slice(0, run.marker.length)) return run.marker.widthPx;
		return runTextWidth(token.text, run, measurer, fonts[token.runIndex]);
	};
	const hangingStopPx = firstLineExtraPx < 0 ? leftPx : undefined;
	let minPx = 0;
	let maxPx = 0;
	let line = Math.max(0, firstLineExtraPx);
	let trailingSpace = 0;
	let chunk = 0;
	let previous: BreakToken | undefined;
	const endLine = () => {
		maxPx = Math.max(maxPx, line - trailingSpace);
		line = 0;
		trailingSpace = 0;
	};
	for (const token of tokens) {
		const width = widthOf(token);
		if (token.kind === 'word' || token.kind === 'object') {
			chunk = previous && !breaksBetween(previous, token) ? chunk + width : width;
			minPx = Math.max(minPx, chunk);
			line += width;
			trailingSpace = 0;
		} else if (token.kind === 'space') {
			line += width;
			trailingSpace += width;
		} else if (token.kind === 'tab') {
			line += placeTab(line, paragraph.tabStops ?? [], hangingStopPx, 0, 0).widthPx;
			trailingSpace = 0;
		} else endLine();
		previous = token.kind === 'word' || token.kind === 'object' ? token : undefined;
	}
	endLine();
	const indents = leftPx + rightPx;
	return { minPx: minPx + indents, maxPx: maxPx + indents };
}

/** The widest range among a cell's blocks; a nested table counts its own column ranges. */
function blocksRange(
	blocks: readonly LayoutBlock[],
	measurer: TextMeasurer,
	nestedTable: (table: Extract<LayoutBlock, { kind: 'table' }>) => ContentRange,
): ContentRange {
	let minPx = 0;
	let maxPx = 0;
	for (const block of blocks) {
		const range = block.kind === 'paragraph' ? paragraphRange(block, measurer) : nestedTable(block);
		minPx = Math.max(minPx, range.minPx);
		maxPx = Math.max(maxPx, range.maxPx);
	}
	return { minPx, maxPx };
}

/** A cell's minimum and maximum widths with its left and right margins, in whole twips. */
export function cellRange(
	cell: LayoutTableCell,
	measurer: TextMeasurer,
	nestedTable: (table: Extract<LayoutBlock, { kind: 'table' }>) => ContentRange,
): ContentRange {
	const content = blocksRange(cell.blocks ?? cell.paragraphs, measurer, nestedTable);
	const margins = (cell.padding?.left ?? 0) + (cell.padding?.right ?? 0);
	const maxPx = content.maxPx > 0 ? content.maxPx : EMPTY_CONTENT_PX;
	const minPx = content.minPx > 0 ? content.minPx : Math.min(maxPx, EMPTY_CONTENT_PX);
	return { minPx: ceilTwips(minPx + margins), maxPx: ceilTwips(Math.max(minPx, maxPx) + margins) };
}
