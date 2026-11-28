import type { TextMeasurer } from './measure';
import { gluedCarry, tokenPosition, type BreakToken, type TokenPosition } from './text-breaks';
import { findHyphenSplit, lineUsedWidthPx } from './hyphenate-line';
import { fittingWordPrefix } from './word-overflow';
import { paragraphShadowInsets } from './paragraph-shadow';
import type { LayoutFragment, LayoutLine, LayoutParagraphFrame } from './result';
import type { LayoutParagraph } from './input';
import { NO_TWIPS, ptToPx, twipsToPx } from './units';
import { legacyTextStartPx } from './legacy-marker';
import { contentControlTagMetrics } from './content-control-tag-metrics';
import { placePositionalTab, placeTab } from './tab-stops';
import {
	fontOf,
	runTextWidth,
	resolveIndents,
	tokenizeParagraph,
	hangsPastLineEnd,
	sourceLengthOf,
	type PlacedToken,
} from './paragraph-tokens';
import { lineMetrics, tokenExtent } from './line-metrics';
import { alignFragments, buildFragments } from './paragraph-fragments';
import { expectDefined } from '../expect-defined';
import { cachedParagraphLayout } from './paragraph-cache';

// Word 16 mode 15 probes in Courier New, Times New Roman, Arial and Arial Narrow.
// The line can give up a quarter of its spaces, with a smaller budget for a short next word.
// Word 16 sweeps in Century Gothic (issue #369) keep "a" and "at" at 0.364 of word plus space and
// drop them by 0.383, so the word part sits between those.
const SPACE_SHRINK = 0.25;
const WORD_SHRINK = 0.375;

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
	startToken: TokenPosition = 0,
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
	startToken: TokenPosition = 0,
): ParagraphLayoutResult {
	const { leftPx, rightPx, firstLineExtraPx: initialIndent } = resolveIndents(paragraph);
	const firstLineExtraPx = startToken ? 0 : initialIndent;
	const bodyWidth = Math.max(1, contentWidthPx - leftPx - rightPx);
	const fonts = paragraph.runs.map(fontOf);
	const { tokens, runOffsets } = tokenizeParagraph(paragraph);
	const startIndex = typeof startToken === 'number' ? startToken : startToken.index;
	const originalWordStarts = new Map<number, number>();
	const resumeWord = (index: number, offset: number) => {
		const word = tokens[index];
		if (word?.kind !== 'word') return;
		const { glue: _glue, ...rest } = word;
		originalWordStarts.set(index, originalWordStarts.get(index) ?? word.sourceStart);
		tokens[index] = {
			...rest,
			sourceStart: word.sourceStart + offset,
			text: word.text.slice(offset),
		};
	};
	if (typeof startToken !== 'number') resumeWord(startToken.index, startToken.offset);
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

	let placed: PlacedToken[] = [];
	let lineWidthPx = 0;
	// Width of the spaces on the current line: justified text may squeeze them to fit one more word.
	let lineSpacePx = 0;
	// Where the current line continues past a picture, and how far its later fragments move for it.
	let jump: { index: number; shiftPx: number } | undefined;
	// How far a line holding a positional tab relative to the margins may reach: to the margin.
	let marginReachPx = 0;
	let consecutiveHyphens = 0;
	let tokenCursor: TokenPosition = startToken;
	let firstToken: TokenPosition = startToken;
	const wordPosition = (index: number, token: BreakToken, advance = 0): TokenPosition => {
		const start = originalWordStarts.get(index) ?? token.sourceStart;
		const offset = token.sourceStart - start + advance;
		return tokenPosition(index, offset);
	};

	// Per-line boxes around wrapped floats, computed when each line starts (its top is then known).
	const boxes: {
		leftInsetPx: number;
		rightInsetPx: number;
		gapBeforePx: number;
		split?: boolean;
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
	/**
	 * Where a line's text starts, from the column's left edge: the paragraph's indent (plus the
	 * first-line offset) or the edge of a wrapped float, whichever is further in. The indent is
	 * not added to the float's edge: Word measures it from the margin.
	 */
	const lineStart = (lineIndex: number) => {
		const indented = leftPx + (lineIndex === 0 ? firstLineExtraPx : 0);
		const inset = boxFor(lineIndex).leftInsetPx;
		// Only beside a float does its edge limit the indent; negative indents stay in the margin.
		return inset > 0 ? Math.max(indented, inset) : indented;
	};
	const lineEnd = (lineIndex: number) => {
		const inset = boxFor(lineIndex).rightInsetPx;
		return inset > 0 ? Math.max(rightPx, inset) : rightPx;
	};
	const availableWidth = (lineIndex: number) =>
		Math.max(1, contentWidthPx - lineStart(lineIndex) - lineEnd(lineIndex));
	const tokenWidth = (token: BreakToken): number => {
		const run = token.kind === 'word' ? paragraph.runs[token.runIndex] : undefined;
		const font = fonts[token.runIndex];
		if (run?.contentControlTag) {
			if (run.contentControlTag === 'end' && !run.blockContentControlTag)
				return ptToPx(contentControlTagMetrics.inlineEndWidthPt);
			return Math.max(
				ptToPx(contentControlTagMetrics.minLabelWidthPt),
				runTextWidth(run.text, run, measurer, font) +
					ptToPx(contentControlTagMetrics.labelPaddingPt),
			);
		}
		// A picture bullet takes its picture's width, not the fallback glyph's.
		if (run?.marker?.widthPx !== undefined && token.kind === 'word' && token.sourceStart === 0)
			if (token.text === run.text.slice(0, run.marker.length)) return run.marker.widthPx;
		return token.kind === 'word' || token.kind === 'space'
			? runTextWidth(
					token.text,
					expectDefined(paragraph.runs[token.runIndex], 'token run'),
					measurer,
					font,
				)
			: token.kind === 'object'
				? (paragraph.runs[token.runIndex]?.object?.widthPx ?? 0)
				: 0;
	};
	const widths = tokens.map(tokenWidth);
	const ends = new Set(['tab', 'lineBreak', 'pageBreak', 'columnBreak']);
	const beforeStart = tokens[startIndex - 1]?.kind;
	let keepLeadingSpaces =
		startIndex === 0 ||
		beforeStart === 'lineBreak' ||
		beforeStart === 'pageBreak' ||
		beforeStart === 'columnBreak';
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
	const placedTokens = () => placed.map((entry) => entry.token);
	// Width of the punctuation hanging past each line's end, which alignment does not count.
	const hangingPx = new Map<number, number>();
	const squeezedLines = new Set<number>();
	const softBreakLines = paragraph.doNotExpandShiftReturn ? new Set<number>() : undefined;

	function flushLine(
		forceEmpty: boolean,
		forcedByBreak: 'page' | 'column' | undefined,
		endedWithAutoHyphen = false,
	) {
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
		const { heightPx, baselinePx, extraBelowPx } = lineMetrics(placed, fonts, paragraph, measurer);
		// Each fragment's box is placed so its baseline sits on the line's.
		rendered.forEach(({ token }, index) => {
			const extent = tokenExtent(token, paragraph, fonts, measurer);
			const fragment = expectDefined(fragments[index], 'fragment for rendered token');
			fragment.sourceStart = globalOffset(token);
			fragment.sourceEnd = globalOffset(
				token,
				token.kind === 'word' || token.kind === 'space' ? token.text.length : 1,
			);
			fragment.topPx = baselinePx - (fragment.contentControlTag ? extent.ascent : extent.above);
			if (token.kind !== 'object') fragment.boxHeightPx = extent.ascent + extent.descent;
		});
		if (jump) {
			boxFor(lineIndex).split = true;
			for (const fragment of fragments.slice(jump.index)) fragment.xPx += jump.shiftPx;
		}
		lines.push({
			yPx: 0,
			heightPx,
			baselinePx,
			fragments,
			sourceStart: start,
			sourceEnd: end,
			...(typeof tokenCursor === 'number'
				? { nextToken: tokenCursor }
				: { nextToken: tokenCursor.index, nextTokenOffset: tokenCursor.offset }),
			...(typeof firstToken === 'number'
				? { firstToken }
				: { firstToken: firstToken.index, firstTokenOffset: firstToken.offset }),
			...(gapBeforePx ? { gapBeforePx } : {}),
			...(extraBelowPx ? { extraBelowPx } : {}),
		});
		runningY += gapBeforePx + heightPx;
		if (forcedByBreak === 'page') pageBreakAfterLine.add(lineIndex);
		if (forcedByBreak === 'column') columnBreakAfterLine.add(lineIndex);
		placed = [];
		lineWidthPx = 0;
		lineSpacePx = 0;
		keepLeadingSpaces = false;
		jump = undefined;
		marginReachPx = 0;
		consecutiveHyphens = endedWithAutoHyphen ? consecutiveHyphens + 1 : 0;
	}

	function tryHyphenate(
		index: number,
		wordToken: Extract<BreakToken, { kind: 'word' }>,
		availablePx: number,
		gapPx: number,
	): boolean {
		const run = paragraph.runs[wordToken.runIndex];
		if (wordToken.glue || run?.synthetic || run?.marker) return false;
		const split = findHyphenSplit(
			wordToken.text,
			availablePx,
			gapPx,
			paragraph.hyphenation,
			consecutiveHyphens,
			(text) => tokenWidth({ ...wordToken, text }),
		);
		if (!split) return false;
		if (!placed.length) firstToken = wordPosition(index, wordToken);
		placed.push({
			token: { ...wordToken, text: split.prefixSource },
			widthPx: split.widthPx,
			displayText: split.displayText,
		});
		tokenCursor = wordPosition(index, wordToken, split.resumeOffset);
		flushLine(false, undefined, split.automatic);
		resumeWord(index, split.resumeOffset);
		widths[index] = tokenWidth(expectDefined(tokens[index], 'resumed word token'));
		return true;
	}

	function placeWordPrefix(
		index: number,
		token: Extract<BreakToken, { kind: 'word' }>,
		prefix: string,
		widthPx: number,
	) {
		if (!placed.length) firstToken = wordPosition(index, token);
		placed.push({ token: { ...token, text: prefix }, widthPx });
		tokenCursor = wordPosition(index, token, prefix.length);
		flushLine(false, undefined);
		resumeWord(index, prefix.length);
		// Fit the remainder's prefix on the next line without shaping all of it again.
		widths[index] = Infinity;
	}

	// Words split across runs (pieces glued to the one before): each piece's first piece, and the
	// width of each whole word by its first piece, measured once.
	const wordStart: number[] = [];
	const wordWidth = new Map<number, number>();
	tokens.forEach((token, index) => {
		const previous = tokens[index - 1];
		const start =
			token.kind === 'word' && token.glue && previous?.kind === 'word'
				? (wordStart[index - 1] ?? index)
				: index;
		wordStart[index] = start;
		const width = widths[index] ?? 0;
		wordWidth.set(start, (wordWidth.get(start) ?? 0) + (Number.isFinite(width) ? width : 0));
	});
	const pieceWidths = [...widths];
	/**
	 * Advance of the word `tokens[index]` belongs to, with the space before it: the squeeze budget of
	 * a word split across runs is the whole word's, so a run boundary inside it never moves a break.
	 */
	function wholeWordPx(index: number, advance: number): number {
		const start = wordStart[index] ?? index;
		const own = pieceWidths[index] ?? 0;
		const total = (wordWidth.get(start) ?? 0) - (Number.isFinite(own) ? own : 0) + advance;
		// The pieces before this one are the last ones placed when the word started on this line.
		const back = index - start;
		const before = back < placed.length ? placed[placed.length - back - 1] : undefined;
		return total + (before?.token.kind === 'space' ? before.widthPx : 0);
	}

	for (let tokenIndex = startIndex; tokenIndex < tokens.length; tokenIndex++) {
		const token = expectDefined(tokens[tokenIndex], 'token index');
		tokenCursor = tokenIndex + 1;
		if (token.kind === 'lineBreak') {
			if (!placed.length) firstToken = tokenIndex;
			softBreakLines?.add(lines.length);
			flushLine(true, undefined);
			keepLeadingSpaces = true;
			continue;
		}
		if (token.kind === 'pageBreak' || token.kind === 'columnBreak') {
			if (!placed.length) firstToken = tokenIndex;
			flushLine(true, token.kind === 'pageBreak' ? 'page' : 'column');
			keepLeadingSpaces = true;
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
		let advance = width * (1 - markerFactor);
		// Automatic East Asian/Western spacing follows the word; it never decides whether it fits.
		const spacePx = token.kind === 'word' ? (token.spaceAfterPx ?? 0) : 0;
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
		} else if (token.kind === 'tab' && paragraph.runs[token.runIndex]?.marker?.legacy) {
			// The text after a legacy marker starts at its legacy indent, not at a tab stop.
			const legacy = expectDefined(paragraph.runs[token.runIndex]?.marker?.legacy, 'legacy');
			const here = lineStart(lines.length) + lineWidthPx;
			width = Math.max(
				0,
				legacyTextStartPx(lineStart(lines.length), here, legacy) - here,
			);
		} else if (token.kind === 'tab') {
			const { followingPx, beforeDecimalPx } = segmentAfter(tokenIndex);
			const tab = placeTab(
				lineStart(lines.length) + lineWidthPx,
				paragraph.tabStops ?? [],
				hangingStopPx,
				followingPx,
				beforeDecimalPx,
				paragraph.defaultTabStopPx,
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
			effectiveAlign === 'justify' &&
			paragraph.squeezeJustifiedSpaces !== false &&
			token.kind === 'word' &&
			lineSpacePx > 0
				? Math.min(lineSpacePx * SPACE_SHRINK, wholeWordPx(tokenIndex, advance) * WORD_SHRINK)
				: 0;
		// A nearly fitting final line of a justified text box should not turn into
		// a stretched nonfinal line. Leave normal word-break decisions alone.
		const wordTolerance =
			effectiveAlign === 'justify' &&
			token.kind === 'word' &&
			!hole &&
			tokenIndex === tokens.length - 1
				? (paragraph.wordFitTolerancePx ?? 0)
				: 0;
		const fits =
			lineWidthPx + (token.kind === 'tab' ? width : advance) <=
			limit + Math.max(squeezePx, wordTolerance);
		const wouldFit =
			fits || placed.every((placed) => paragraph.runs[placed.token.runIndex]?.synthetic);
		// A word that no longer fits before a picture continues after it, on the same line.
		const pastHole =
			!fits && hole && !jump && token.kind === 'word'
				? hole.endPx - lineStart(lineIndex) - lineWidthPx
				: undefined;
		// Line-end punctuation that may not start the next line hangs past the margin instead.
		const hanging =
			!wouldFit &&
			!hole &&
			!jump &&
			placed.length > 0 &&
			hangsPastLineEnd(paragraph, tokens, tokenIndex);
		if (hanging) hangingPx.set(lineIndex, advance);
		if (pastHole !== undefined && lineWidthPx + pastHole + advance <= availableWidth(lineIndex)) {
			jump = { index: placed.length, shiftPx: pastHole };
		} else if (!wouldFit && token.kind !== 'space' && !hanging) {
			if (token.kind === 'word' && !jump) {
				const availableForWord = limit + squeezePx - lineWidthPx;
				const gap = limit - lineUsedWidthPx(placed, lineWidthPx);
				if (tryHyphenate(tokenIndex, token, availableForWord, gap)) {
					tokenIndex--;
					continue;
				}
			}
			// Kinsoku and words split across runs: the glued tokens before this one go down with it.
			const carry = jump ? 0 : gluedCarry(tokens, tokenIndex, placedTokens());
			if (carry) {
				placed.splice(placed.length - carry);
				tokenCursor = tokenIndex - carry;
				flushLine(false, undefined);
				tokenIndex -= carry + 1;
				continue;
			}
			// An overlong word already on its own line can break inside its next run.
			if (
				!jump &&
				token.kind === 'word' &&
				token.glue === 'word' &&
				placed.at(-1)?.token === tokens[tokenIndex - 1]
			) {
				const available = limit - lineWidthPx;
				const measure = (text: string) => tokenWidth({ ...token, text });
				const prefix = fittingWordPrefix(token.text, available, measure);
				const prefixWidth = measure(prefix);
				if (prefix.length < token.text.length && prefixWidth <= available) {
					placeWordPrefix(tokenIndex, token, prefix, prefixWidth);
					tokenIndex--;
					continue;
				}
			}
			// On an otherwise unbreakable line, keep the control tag attached even if it overhangs.
			if (token.glue !== 'controlTag') {
				tokenCursor = tokenIndex;
				if (token.kind === 'word') tokenCursor = wordPosition(tokenIndex, token);
				flushLine(false, undefined);
				tokenIndex--;
				continue;
			}
		}
		if (
			token.kind === 'word' &&
			!marker &&
			!paragraph.runs[token.runIndex]?.synthetic &&
			!hanging
		) {
			const available = jump
				? availableWidth(lineIndex) - jump.shiftPx - lineWidthPx
				: limit - lineWidthPx;
			if (
				width > available + wordTolerance &&
				placed.every(({ token }) => paragraph.runs[token.runIndex]?.synthetic)
			) {
				if (
					!jump &&
					tryHyphenate(tokenIndex, token, available + wordTolerance, available + wordTolerance)
				) {
					tokenIndex--;
					continue;
				}
				const prefix = fittingWordPrefix(token.text, available + wordTolerance, (text) =>
					tokenWidth({ ...token, text }),
				);
				if (prefix.length < token.text.length) {
					placeWordPrefix(tokenIndex, token, prefix, tokenWidth({ ...token, text: prefix }));
					tokenIndex--;
					continue;
				}
				width = advance = tokenWidth(token);
				widths[tokenIndex] = width;
			}
		}
		if (token.kind === 'space' && placed.length === 0 && !keepLeadingSpaces) continue;
		if (!placed.length) firstToken = wordPosition(tokenIndex, token);
		placed.push({
			token,
			widthPx: width,
			...(leader ? { leader } : {}),
			...(markerFactor ? { xOffsetPx: -width * markerFactor } : {}),
			...(markerFactor || spacePx ? { advancePx: advance + spacePx } : {}),
		});
		if (fits && squeezePx > wordTolerance && lineWidthPx + advance > limit + wordTolerance)
			squeezedLines.add(lineIndex);
		lineWidthPx += token.kind === 'tab' ? width : advance + spacePx;
		if (token.kind === 'space') lineSpacePx += advance;
	}
	// A cell's terminal paragraph with no content consists solely of its end mark.
	const hiddenCellMark = paragraph.hideCellEndMark && tokens.length === 0;
	if (!hiddenCellMark) flushLine(lines.length === 0, undefined);

	const lastIndex = lines.length - 1;
	for (const [index, line] of lines.entries()) {
		const hole = boxes[index]?.hole;
		const width = hole
			? Math.min(availableWidth(index), hole.startPx - lineStart(index))
			: availableWidth(index);
		lines[index] = {
			...line,
			fragments: alignFragments(
				line.fragments,
				width + (hangingPx.get(index) ?? 0),
				// Measured in Word 16: a line ending in a page or column break keeps its natural
				// width with or without doNotExpandShiftReturn.
				index === lastIndex ||
					softBreakLines?.has(index) === true ||
					pageBreakAfterLine.has(index) ||
					columnBreakAfterLine.has(index),
				boxes[index]?.split ? 'left' : effectiveAlign,
				squeezedLines.has(index),
			).map((fragment) => ({ ...fragment, xPx: fragment.xPx + lineStart(index) })),
		};
	}

	// Borders sit outside the text: top and bottom lines (with their gaps) add to the height, and
	// side lines are drawn beyond the indents.
	const borders = paragraph.borders;
	const extent = (side: keyof NonNullable<typeof borders>) =>
		borders?.[side] ? borders[side]!.widthPx + borders[side]!.spacePx : 0;
	const shadow = paragraphShadowInsets(borders);
	const insetTopPx = extent('top') + shadow.topPx;
	const insetBottomPx = extent('bottom') + shadow.bottomPx;
	let y = insetTopPx;
	for (const line of lines) {
		y += line.gapBeforePx ?? 0;
		line.yPx = y;
		y += line.heightPx;
	}
	const hanging = borders ? Math.min(0, firstLineExtraPx) : 0;
	const rtl = paragraph.direction === 'rtl';
	const frame =
		borders || paragraph.shading
			? {
					leftPx: leftPx - extent('left') + (rtl ? 0 : hanging),
					widthPx: bodyWidth + extent('left') + extent('right') - hanging,
					...(borders ? { borders } : {}),
					...(shadow.topPx ? { shadowTopPx: shadow.topPx } : {}),
					...(shadow.bottomPx ? { shadowBottomPx: shadow.bottomPx } : {}),
					...(paragraph.shading ? { shading: paragraph.shading } : {}),
				}
			: undefined;

	return {
		lines,
		spacingBeforePx: hiddenCellMark ? 0 : twipsToPx(paragraph.spacingBeforeTwips ?? NO_TWIPS),
		spacingAfterPx: hiddenCellMark ? 0 : twipsToPx(paragraph.spacingAfterTwips ?? NO_TWIPS),
		contentHeightPx: y + insetBottomPx,
		insetTopPx,
		insetBottomPx,
		...(frame ? { frame } : {}),
		pageBreakAfterLine,
		columnBreakAfterLine,
	};
}
