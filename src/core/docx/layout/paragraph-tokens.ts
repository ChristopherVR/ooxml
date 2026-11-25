import { definedProps } from '../defined-props';
import type { LayoutFontSpec, TextMeasurer } from './measure';
import type { LayoutFragment } from './result';
import { hangsAtLineEnd, kinsokuSets, type KinsokuSets } from '../../text/kinsoku';
import { AUTO_SPACE_EM, eastAsianAutoSpace } from '../../text/east-asian';
import { stripSoftHyphens } from '../../text/hyphenation';
import {
	appendBreakMarker,
	breakAllowedBetween,
	graphemesOf,
	kinsokuForbids,
	tokenizeRun,
	type BreakToken,
} from './text-breaks';
import type { LayoutParagraph, LayoutRun } from './input';
import { NO_TWIPS, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units';

/** Word draws superscript and subscript text at about two thirds of the run's size. */
export const SCRIPT_SCALE = 0.65;

export function fontOf(run: LayoutRun): LayoutFontSpec {
	return {
		family: run.fontFamily || DEFAULT_FONT_FAMILY,
		sizePx: ptToPx(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (run.script ? SCRIPT_SCALE : 1),
		...definedProps({
			bold: run.bold,
			italic: run.italic,
			ligatures: run.ligatures,
			smallCaps: run.smallCaps,
			fallback: run.fontFallback,
		}),
		...(run.kerningThresholdPt === undefined
			? {}
			: {
					kerning:
						run.kerningThresholdPt > 0 &&
						(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT) >= run.kerningThresholdPt
							? ('normal' as const)
							: ('none' as const),
				}),
	};
}

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
/** Word applies additional character spacing after scaling the glyph advances. */
export function runTextWidth(
	text: string,
	run: LayoutRun,
	measurer: TextMeasurer,
	font = fontOf(run),
): number {
	const visible = stripSoftHyphens(text);
	let characters = 0;
	if (run.characterSpacingPx) {
		for (const _segment of graphemes.segment(visible)) characters++;
	}
	return Math.max(
		0,
		(measurer.widthOf(visible, font) * (run.textScalePercent ?? 100)) / 100 +
			characters * (run.characterSpacingPx ?? 0),
	);
}

export function resolveIndents(paragraph: LayoutParagraph): {
	leftPx: number;
	rightPx: number;
	firstLineExtraPx: number;
} {
	const rtl = paragraph.direction === 'rtl';
	const startTwips =
		paragraph.indentStartTwips ??
		(rtl ? paragraph.indentRightTwips : paragraph.indentLeftTwips) ??
		NO_TWIPS;
	const endTwips =
		paragraph.indentEndTwips ??
		(rtl ? paragraph.indentLeftTwips : paragraph.indentRightTwips) ??
		NO_TWIPS;
	const leftTwips = rtl ? endTwips : startTwips;
	const rightTwips = rtl ? startTwips : endTwips;
	const firstLineExtraPx =
		paragraph.hangingTwips != null
			? -twipsToPx(paragraph.hangingTwips)
			: twipsToPx(paragraph.firstLineTwips ?? NO_TWIPS);
	return {
		leftPx: twipsToPx(leftTwips),
		rightPx: twipsToPx(rightTwips),
		firstLineExtraPx,
	};
}

/** Source atoms occupy one editor position even when they have no visible text. */
export function sourceLengthOf(run: LayoutRun): number {
	return run.sourceLength ?? (run.object ? 1 : run.text.length || (run.breakAfter ? 1 : 0));
}

/** The kinsoku sets for a run's characters, or undefined when kinsoku does not apply to it. */
function runKinsoku(
	paragraph: LayoutParagraph,
	run: LayoutRun | undefined,
): KinsokuSets | undefined {
	if (paragraph.kinsoku === false || !run?.kinsokuLanguage || run.synthetic) return undefined;
	return kinsokuSets(run.kinsokuLanguage, paragraph.kinsokuSettings);
}

/**
 * Marks the first word of a run that continues the previous run's word: glued when the boundary
 * between the two runs is no break opportunity, or kinsoku forbids it.
 */
function glueAcrossRuns(paragraph: LayoutParagraph, tokens: BreakToken[], start: number): void {
	const token = tokens[start];
	const previous = tokens[start - 1];
	if (token?.kind !== 'word' || previous?.kind !== 'word') return;
	const before = paragraph.runs[previous.runIndex];
	const after = paragraph.runs[token.runIndex];
	if (before?.synthetic || before?.marker) return;
	const last = graphemesOf(previous.text).at(-1) ?? '';
	const first = graphemesOf(token.text)[0] ?? '';
	if (!breakAllowedBetween(last, first)) token.glue = 'word';
	else if (kinsokuForbids(last, first, runKinsoku(paragraph, before), runKinsoku(paragraph, after)))
		token.glue = 'kinsoku';
}

/**
 * Word's automatic space between East Asian letters and Western letters or digits, a quarter em
 * of the preceding character, recorded on the word before each such boundary.
 */
function applyAutoSpace(paragraph: LayoutParagraph, tokens: BreakToken[]): void {
	if (paragraph.autoSpaceDE === false && paragraph.autoSpaceDN === false) return;
	for (let index = 1; index < tokens.length; index++) {
		const previous = tokens[index - 1];
		const token = tokens[index];
		if (previous?.kind !== 'word' || token?.kind !== 'word') continue;
		const run = paragraph.runs[previous.runIndex];
		if (!run || run.synthetic || paragraph.runs[token.runIndex]?.synthetic) continue;
		const kind = eastAsianAutoSpace(
			graphemesOf(previous.text).at(-1) ?? '',
			graphemesOf(token.text)[0] ?? '',
		);
		if (!kind || (kind === 'alpha' ? paragraph.autoSpaceDE : paragraph.autoSpaceDN) === false)
			continue;
		previous.spaceAfterPx = fontOf(run).sizePx * AUTO_SPACE_EM;
	}
}

/**
 * Whether `tokens[index]`, which does not fit on its line, hangs past the line end instead
 * (`w:overflowPunct`): a single kinsoku-glued mark of the hanging set in a justified paragraph,
 * not followed by another character that may not start a line.
 */
export function hangsPastLineEnd(
	paragraph: LayoutParagraph,
	tokens: readonly BreakToken[],
	index: number,
): boolean {
	const token = tokens[index];
	if (token?.kind !== 'word' || token.glue !== 'kinsoku') return false;
	if (paragraph.overflowPunct === false || paragraph.align !== 'justify') return false;
	const language = paragraph.runs[token.runIndex]?.kinsokuLanguage;
	const [mark, ...rest] = graphemesOf(token.text);
	if (!language || !mark || rest.length || !hangsAtLineEnd(mark, language)) return false;
	const next = tokens[index + 1];
	return !(next?.kind === 'word' && next.glue);
}

export function tokenizeParagraph(paragraph: LayoutParagraph): {
	tokens: BreakToken[];
	runOffsets: number[];
} {
	const tokens: BreakToken[] = [];
	const runOffsets: number[] = [];
	let cursor = 0;
	paragraph.runs.forEach((run, runIndex) => {
		const start = tokens.length;
		runOffsets.push(cursor);
		if (run.hidden) {
			if (!run.synthetic) cursor += sourceLengthOf(run);
			return;
		}
		if (run.object) tokens.push({ kind: 'object', runIndex, sourceStart: 0 });
		else if (run.marker) {
			// A marker such as "Chapter 1.2" is indivisible; its separator still acts
			// as a real tab or space after the aligned glyphs.
			tokens.push({
				kind: 'word',
				runIndex,
				sourceStart: 0,
				text: run.text.slice(0, run.marker.length),
			});
			tokens.push(...tokenizeRun(run.text.slice(run.marker.length), runIndex));
		} else {
			tokens.push(...tokenizeRun(run.text, runIndex, runKinsoku(paragraph, run)));
			glueAcrossRuns(paragraph, tokens, start);
		}
		if (run.breakAfter) appendBreakMarker(tokens, runIndex, run.text.length, run.breakAfter);
		// A list label maps to the start of the paragraph's text for click-to-cursor.
		if (run.synthetic) {
			for (const token of tokens) if (token.runIndex === runIndex) token.sourceStart = 0;
		} else cursor += sourceLengthOf(run);
	});
	applyAutoSpace(paragraph, tokens);
	return { tokens, runOffsets };
}

export interface PlacedToken {
	token: BreakToken;
	widthPx: number;
	/** Rendered text when it differs from `token.text` (e.g. a hyphenated prefix with `-`). */
	displayText?: string;
	/** Glyph position and consumed advance differ for centered/right-aligned list markers. */
	xOffsetPx?: number;
	advancePx?: number;
	leader?: LayoutFragment['leader'];
}
