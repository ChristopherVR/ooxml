import type { TextMeasurer } from './measure';
import type { BreakToken } from './text-breaks';
import type { LayoutFragment, LayoutLine, LayoutParagraphFrame } from './result';
import type { LayoutParagraph } from './input';
import { NO_TWIPS, twipsToPx } from './units';
import { placePositionalTab, placeTab } from './tab-stops';
import {
	fontOf,
	runTextWidth,
	resolveIndents,
	tokenizeParagraph,
	sourceLengthOf,
	type PlacedToken,
} from './paragraph-tokens';
import { lineMetrics, tokenExtent } from './line-metrics';
import { alignFragments, buildFragments } from './paragraph-fragments';
import { expectDefined } from '../expect-defined';
import { cachedParagraphLayout } from './paragraph-cache';

/**
 * How far justified text may squeeze its spaces, as a fraction of their width, to fit one more word.
 * Word 16 does this. The value is fitted to probes of justified Courier New paragraphs from 9 to
 * 13 pt (it reproduces all but 2 of 106 line-end decisions; Word accepted up to 0.20 at 13 pt and
 * refused 0.14 at 10 pt, so no constant is exact), not a documented rule: the one number to revisit
 * if more references disagree.
 */
const SPACE_SHRINK = 0.145;

export interface ParagraphLayoutResult {
	/** Lines with `yPx` relative to the paragraph box's own top (0 for the first line). */
	lines: LayoutLine[];
	spacingBeforePx: number;
	spacingAfterPx: number;
	contentHeightPx: number;
	/** Indices (into `lines`) after which an explicit page break token was consumed. */
	pageBreakAfterLine: Set<number>;
	/** Indices after which an explicit column break token was consumed. */
	columnBreakAfterLine: Set<number>;
	/** Space above the first line and below the last taken by top and bottom borders. */
	insetTopPx: number;
	insetBottomPx: number;
	frame?: LayoutParagraphFrame;
}

/**
 * Lays out one paragraph's lines within `contentWidthPx` (the column/table
 * cell width available to the paragraph, before indents are subtracted).
 * Page and column breaks are reported back as line indices; the page-flow
 * stage decides what to do with them (start a new page/column) since only it
 * knows the current page/column state.
 */
/**
 * Space taken from one line by wrapped floating pictures: insets from the line's left and right
 * (beyond the paragraph indents) and a gap to skip first (text below a picture). `yPx` is the
 * line's top relative to the paragraph's first line.
 */
export type LineBoxFn = (
	yPx: number,
	heightPx: number,
) => {
	leftInsetPx: number;
	rightInsetPx: number;
	gapBeforePx: number;
	/** A picture the line runs around: text fills the space before it, then the space after it. */
	hole?: { startPx: number; endPx: number };
};

export function layoutParagraph(
	paragraph: LayoutParagraph,
	contentWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	lineBox?: LineBoxFn,
	startToken = 0,
): ParagraphLayoutResult {
	if (lineBox)
		return measureParagraph(paragraph, contentWidthPx, measurer, note, lineBox, startToken);
	return cachedParagraphLayout(
		measurer,
		() => ({
			index: JSON.stringify([paragraph.id, contentWidthPx, startToken]),
			signature: JSON.stringify(paragraph),
		}),
		note,
		(report) =>
			measureParagraph(paragraph, contentWidthPx, measurer, report, undefined, startToken),
	);
}

function measureParagraph(
	paragraph: LayoutParagraph,
	contentWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	lineBox?: LineBoxFn,
	startToken = 0,
): ParagraphLayoutResult {
	const { leftPx, rightPx, firstLineExtraPx: initialIndent } = resolveIndents(paragraph);
	const firstLineExtraPx = startToken ? 0 : initialIndent;
	const bodyWidth = Math.max(1, contentWidthPx - leftPx - rightPx);
	const fonts = paragraph.runs.map(fontOf);
	const { tokens, runOffsets } = tokenizeParagraph(paragraph);
	const globalOffset = (token: BreakToken, advance = 0) => {
		const run = expectDefined(paragraph.runs[token.runIndex], 'source run');
		const length = sourceLengthOf(run);
		const offset = run.synthetic
			? 0
			: Math.max(
					0,
					Math.min(
						length,
						Math.round(((token.sourceStart + advance) * length) / Math.max(1, run.text.length)),
					),
				);
		return expectDefined(runOffsets[token.runIndex], 'run offset index') + offset;
	};
	const effectiveAlign = paragraph.align ?? (paragraph.direction === 'rtl' ? 'right' : 'left');

	const lines: LayoutLine[] = [];
	const pageBreakAfterLine = new Set<number>();
	const columnBreakAfterLine = new Set<number>();
	let widthOverflowNoted = false;

	let placed: PlacedToken[] = [];
	let lineWidthPx = 0;
	// Width of the spaces on the current line: justified text may squeeze them to fit one more word.
	let lineSpacePx = 0;
	// Where the current line continues past a picture, and how far its later fragments move for it.
	let jump: { index: number; shiftPx: number } | undefined;
	// How far a line holding a positional tab relative to the margins may reach: to the margin.
	let marginReachPx = 0;
	let tokenCursor = startToken;
	let firstToken = startToken;

	// Per-line boxes around wrapped floats, computed when each line starts (its top is then known).
	const boxes: {
		leftInsetPx: number;
		rightInsetPx: number;
		gapBeforePx: number;
		hole?: { startPx: number; endPx: number };
	}[] = [];
	let runningY = 0;
	const guessHeight = measurer.lineHeightOf(fonts[0] ?? fontOf({ text: '' }));
	const boxFor = (lineIndex: number) => {
		if (!boxes[lineIndex]) {
			let gap = 0;
			let box = lineBox?.(runningY, guessHeight);
			// Skipping below one picture can reach another; settle within a few steps.
			for (let step = 0; box && box.gapBeforePx > 0 && step < 8; step++) {
				gap += box.gapBeforePx;
				box = lineBox!(runningY + gap, guessHeight);
			}
			boxes[lineIndex] = {
				leftInsetPx: box?.leftInsetPx ?? 0,
				rightInsetPx: box?.rightInsetPx ?? 0,
				gapBeforePx: gap,
				...(box?.hole ? { hole: box.hole } : {}),
			};
		}
		return expectDefined(boxes[lineIndex], 'line box');
	};
	const availableWidth = (lineIndex: number) => {
		const box = boxFor(lineIndex);
		return Math.max(
			1,
			bodyWidth - (lineIndex === 0 ? firstLineExtraPx : 0) - box.leftInsetPx - box.rightInsetPx,
		);
	};
	/** Where a line's text starts, from the paragraph's text margin (indent plus first-line offset). */
	const lineStart = (lineIndex: number) =>
		leftPx + (lineIndex === 0 ? firstLineExtraPx : 0) + boxFor(lineIndex).leftInsetPx;
	const tokenWidth = (token: BreakToken): number =>
		token.kind === 'word' || token.kind === 'space'
			? runTextWidth(
					token.text,
					expectDefined(paragraph.runs[token.runIndex], 'token run'),
					measurer,
					fonts[token.runIndex],
				)
			: token.kind === 'object'
				? (paragraph.runs[token.runIndex]?.object?.widthPx ?? 0)
				: 0;
	const widths = tokens.map(tokenWidth);
	const ends = new Set(['tab', 'lineBreak', 'pageBreak', 'columnBreak']);
	/** Width of the text after the tab at `index`, whole and up to its first decimal separator. */
	function segmentAfter(index: number): { followingPx: number; beforeDecimalPx: number } {
		let followingPx = 0;
		let beforeDecimalPx: number | undefined;
		for (let next = index + 1; next < tokens.length; next++) {
			const token = expectDefined(tokens[next], 'token index');
			if (ends.has(token.kind)) break;
			if (beforeDecimalPx === undefined && token.kind === 'word' && /[.,]/.test(token.text)) {
				const prefix = token.text.slice(0, token.text.search(/[.,]/));
				beforeDecimalPx =
					followingPx +
					runTextWidth(
						prefix,
						expectDefined(paragraph.runs[token.runIndex], 'decimal run'),
						measurer,
					);
			}
			followingPx += widths[next] ?? 0;
		}
		return { followingPx, beforeDecimalPx: beforeDecimalPx ?? followingPx };
	}
	const hangingStopPx = firstLineExtraPx < 0 ? leftPx : undefined;

	function flushLine(forceEmpty: boolean, forcedByBreak: 'page' | 'column' | undefined) {
		if (!placed.length && !forceEmpty) return;
		const lineIndex = lines.length;
		// Trailing whitespace at a wrap point is consumed (counts toward sourceEnd)
		// but never rendered or measured, so it cannot skew alignment/justification.
		const rendered = [...placed];
		while (rendered.length && rendered.at(-1)!.token.kind === 'space') rendered.pop();
		const fragments = buildFragments(rendered, paragraph);
		const [firstPlaced] = placed;
		// The first line of a paragraph starts at its first character, even when that is an anchor or
		// field code that draws nothing in the line (Word reports offset 0 there too).
		const start =
			lines.length === 0 && startToken === 0
				? 0
				: firstPlaced
					? globalOffset(firstPlaced.token)
					: 0;
		const last = placed.at(-1)?.token;
		const end = last
			? globalOffset(last, last.kind === 'word' || last.kind === 'space' ? last.text.length : 1)
			: start;
		const gapBeforePx = boxFor(lineIndex).gapBeforePx;
		const { heightPx, baselinePx } = lineMetrics(placed, fonts, paragraph, measurer);
		// Each fragment's box is placed so its baseline sits on the line's.
		rendered.forEach(({ token }, index) => {
			const extent = tokenExtent(token, paragraph, fonts, measurer);
			const fragment = expectDefined(fragments[index], 'fragment for rendered token');
			fragment.sourceStart = globalOffset(token);
			fragment.sourceEnd = globalOffset(
				token,
				token.kind === 'word' || token.kind === 'space' ? token.text.length : 1,
			);
			fragment.topPx = baselinePx - extent.above;
			if (token.kind !== 'object') fragment.boxHeightPx = extent.ascent + extent.descent;
		});
		if (jump) {
			for (const fragment of fragments.slice(jump.index)) fragment.xPx += jump.shiftPx;
		}
		lines.push({
			yPx: 0,
			heightPx,
			baselinePx,
			fragments,
			sourceStart: start,
			sourceEnd: end,
			nextToken: tokenCursor,
			firstToken,
			...(gapBeforePx ? { gapBeforePx } : {}),
		});
		runningY += gapBeforePx + heightPx;
		if (forcedByBreak === 'page') pageBreakAfterLine.add(lineIndex);
		if (forcedByBreak === 'column') columnBreakAfterLine.add(lineIndex);
		placed = [];
		lineWidthPx = 0;
		lineSpacePx = 0;
		jump = undefined;
		marginReachPx = 0;
	}

	for (const [tokenIndex, token] of tokens.entries()) {
		if (tokenIndex < startToken) continue;
		tokenCursor = tokenIndex + 1;
		if (token.kind === 'lineBreak') {
			if (!placed.length) firstToken = tokenIndex;
			flushLine(true, undefined);
			continue;
		}
		if (token.kind === 'pageBreak' || token.kind === 'columnBreak') {
			if (!placed.length) firstToken = tokenIndex;
			flushLine(true, token.kind === 'pageBreak' ? 'page' : 'column');
			continue;
		}
		let leader: LayoutFragment['leader'];
		let width = expectDefined(widths[tokenIndex], 'token width');
		const marker = paragraph.runs[token.runIndex]?.marker;
		const markerFactor =
			marker && token.kind === 'word' && tokenIndex === 0
				? marker.alignment === 'right'
					? 1
					: marker.alignment === 'center'
						? 0.5
						: 0
				: 0;
		const advance = width * (1 - markerFactor);
		const positional = token.kind === 'tab' && paragraph.runs[token.runIndex]?.positionalTab;
		if (positional) {
			const { followingPx } = segmentAfter(tokenIndex);
			// Word aligns to the closer of the indent and the margin: never past the margins, and a
			// tab to the left margin of an indented paragraph leaves its text at the indent.
			const [startPx, endPx] =
				positional.relativeTo === 'indent'
					? [Math.max(leftPx, 0), Math.min(contentWidthPx - rightPx, contentWidthPx)]
					: [0, contentWidthPx];
			const placeAt = (positionPx: number) =>
				placePositionalTab(positionPx, positional, startPx, endPx, followingPx);
			let tab = placeAt(lineStart(lines.length) + lineWidthPx);
			if (!tab && placed.length) {
				// The line is past the tab's point: the tab goes to that point on the next line.
				tokenCursor = tokenIndex;
				flushLine(false, undefined);
				tokenCursor = tokenIndex + 1;
				tab = placeAt(lineStart(lines.length));
			}
			width = tab?.widthPx ?? 0;
			leader = tab?.leader === 'none' ? undefined : tab?.leader;
			if (positional.relativeTo === 'margin')
				marginReachPx = contentWidthPx - lineStart(lines.length);
		} else if (token.kind === 'tab') {
			const { followingPx, beforeDecimalPx } = segmentAfter(tokenIndex);
			const tab = placeTab(
				lineStart(lines.length) + lineWidthPx,
				paragraph.tabStops ?? [],
				hangingStopPx,
				followingPx,
				beforeDecimalPx,
			);
			width = tab.widthPx;
			leader = tab.leader === 'none' ? undefined : tab.leader;
		}
		const lineIndex = lines.length;
		const hole = boxFor(lineIndex).hole;
		const beforeHole = hole && !jump ? hole.startPx - lineStart(lineIndex) : undefined;
		const limit = Math.max(
			beforeHole ?? availableWidth(lineIndex) - (jump?.shiftPx ?? 0),
			marginReachPx,
		);
		// A pending trailing space only ever gets counted once it turns out to be
		// internal (i.e. this token fits after it); a trailing space at an actual
		// wrap point is trimmed in flushLine instead, so it must not be
		// pre-subtracted here.
		const squeezePx =
			effectiveAlign === 'justify' && token.kind === 'word' ? lineSpacePx * SPACE_SHRINK : 0;
		const fits = lineWidthPx + (token.kind === 'tab' ? width : advance) <= limit + squeezePx;
		const wouldFit =
			fits || placed.every((placed) => paragraph.runs[placed.token.runIndex]?.synthetic);
		// A word that no longer fits before a picture continues after it, on the same line.
		const pastHole =
			!fits && hole && !jump && token.kind === 'word'
				? hole.endPx - lineStart(lineIndex) - lineWidthPx
				: undefined;
		if (pastHole !== undefined && lineWidthPx + pastHole + advance <= availableWidth(lineIndex)) {
			jump = { index: placed.length, shiftPx: pastHole };
		} else if (!wouldFit && token.kind !== 'space') {
			tokenCursor = tokenIndex;
			flushLine(false, undefined);
			tokenCursor = tokenIndex + 1;
			if (width > availableWidth(lines.length) && !widthOverflowNoted) {
				widthOverflowNoted = true;
				note(
					'A word wider than its line/column was not split mid-word; it overflows the line box.',
				);
			}
		}
		if (token.kind === 'space' && placed.length === 0) continue; // leading space at wrap point is dropped
		if (!placed.length) firstToken = tokenIndex;
		placed.push({
			token,
			widthPx: width,
			...(leader ? { leader } : {}),
			...(markerFactor ? { xOffsetPx: -width * markerFactor, advancePx: advance } : {}),
		});
		lineWidthPx += token.kind === 'tab' ? width : advance;
		if (token.kind === 'space') lineSpacePx += advance;
	}
	flushLine(lines.length === 0, undefined);

	const lastIndex = lines.length - 1;
	for (const [index, line] of lines.entries()) {
		lines[index] = {
			...line,
			fragments: alignFragments(
				line.fragments,
				availableWidth(index),
				index === lastIndex,
				// A line that runs around a picture is not aligned across it.
				boxes[index]?.hole ? 'left' : effectiveAlign,
			).map((fragment) => ({ ...fragment, xPx: fragment.xPx + lineStart(index) })),
		};
	}

	// Borders sit outside the text: top and bottom lines (with their gaps) add to the height, and
	// side lines are drawn beyond the indents.
	const borders = paragraph.borders;
	const extent = (side: keyof NonNullable<typeof borders>) =>
		borders?.[side] ? borders[side]!.widthPx + borders[side]!.spacePx : 0;
	const insetTopPx = extent('top');
	const insetBottomPx = extent('bottom');
	let y = insetTopPx;
	for (const line of lines) {
		y += line.gapBeforePx ?? 0;
		line.yPx = y;
		y += line.heightPx;
	}
	const frame =
		borders || paragraph.shading
			? {
					leftPx: leftPx - extent('left'),
					widthPx: bodyWidth + extent('left') + extent('right'),
					...(borders ? { borders } : {}),
					...(paragraph.shading ? { shading: paragraph.shading } : {}),
				}
			: undefined;

	return {
		lines,
		spacingBeforePx: twipsToPx(paragraph.spacingBeforeTwips ?? NO_TWIPS),
		spacingAfterPx: twipsToPx(paragraph.spacingAfterTwips ?? NO_TWIPS),
		contentHeightPx: y + insetBottomPx,
		insetTopPx,
		insetBottomPx,
		...(frame ? { frame } : {}),
		pageBreakAfterLine,
		columnBreakAfterLine,
	};
}
