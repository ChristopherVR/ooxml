import { complexBreakAllowed } from '../../text/complex-breaks';
import { isEastAsianChar } from '../../text/east-asian';
import { prohibitedAtLineEnd, prohibitedAtLineStart, type KinsokuSets } from '../../text/kinsoku';

/**
 * A UAX#14-lite tokenizer: enough of the Unicode line-breaking algorithm to place break
 * opportunities at spaces, after hyphens and around East Asian characters, without the full class
 * table. Kinsoku (the East Asian line-start and line-end prohibition, see `text/kinsoku.ts`) does
 * not remove a break opportunity from the token stream: the token after it is marked `glue` and
 * the line breaker moves it down together with the token before it.
 */
export { isEastAsianChar };

/**
 * Why a token may not start a line on its own: `kinsoku` (a prohibited character at the
 * boundary), `word` (a word split between formatted runs), `controlTag` (a tag and its content), or `nbsp` (a space that touches a non-breaking space,
 * and the word after it).
 */
export type TokenGlue = 'kinsoku' | 'word' | 'controlTag' | 'nbsp';

/** Resume at a token boundary, or within a word split across lines. */
export type TokenPosition = number | { index: number; offset: number };

export function tokenPosition(index: number, offset?: number): TokenPosition {
	return offset ? { index, offset } : index;
}

export type BreakToken = { runIndex: number; sourceStart: number; glue?: TokenGlue } & (
	| {
			kind: 'word';
			text: string;
			/** Automatic East Asian/Western spacing after this word (`w:autoSpaceDE`/`DN`), in px. */
			spaceAfterPx?: number;
	  }
	| { kind: 'space'; text: string }
	| { kind: 'tab' }
	/** An inline picture: an unbreakable box of the run's object size. */
	| { kind: 'object' }
	/** A manual line break (`w:br`/`w:cr`); ends the current line without ending the paragraph. */
	| { kind: 'lineBreak' }
	/** An explicit page or column break carried on a run (see {@link LayoutRun.breakAfter}). */
	| { kind: 'pageBreak' | 'columnBreak' }
);

const ZERO_WIDTH_SPACE = '​';

/**
 * Whether the line breaker may break between two adjacent graphemes of text, before kinsoku: after
 * a plain hyphen, an en or em dash, or a zero width space, before or after an East Asian
 * character, and between graphemes of Thai, Lao, Khmer or Myanmar (`text/complex-breaks.ts`). U+2011
 * (non-breaking hyphen, `w:noBreakHyphen`) is never split on.
 */
export function breakAllowedBetween(before: string, after: string): boolean {
	if (!before || !after) return false;
	if (
		before === '-' ||
		before === '\u2013' ||
		before === '\u2014' ||
		before.endsWith(ZERO_WIDTH_SPACE)
	)
		return true;
	return isEastAsianChar(before) || isEastAsianChar(after) || complexBreakAllowed(before, after);
}

/** Whether kinsoku forbids a break between `before` (under `beforeSets`) and `after`. */
export function kinsokuForbids(
	before: string,
	after: string,
	beforeSets: KinsokuSets | undefined,
	afterSets: KinsokuSets | undefined,
): boolean {
	return (
		(afterSets !== undefined && prohibitedAtLineStart(after, afterSets)) ||
		(beforeSets !== undefined && prohibitedAtLineEnd(before, beforeSets))
	);
}

const graphemeSegmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

/** The grapheme clusters of `text`. */
export function graphemesOf(text: string): string[] {
	return [...graphemeSegmenter.segment(text)].map((part) => part.segment);
}

/**
 * Splits one run's text into break-opportunity tokens. `sourceStart` is the
 * character offset of each token within the run's own text, so callers can
 * map a line's fragments back to a paragraph/run character range (used for
 * click-to-place-cursor mapping in the Print Layout view). `kinsoku` is the
 * run's prohibition sets, when kinsoku applies to it.
 */
export function tokenizeRun(text: string, runIndex: number, kinsoku?: KinsokuSets): BreakToken[] {
	const tokens: BreakToken[] = [];
	let word = '';
	let wordStart = 0;
	let glue: TokenGlue | undefined;
	let previous = '';
	const flushWord = () => {
		if (word)
			tokens.push({
				kind: 'word',
				text: word,
				runIndex,
				sourceStart: wordStart,
				...(glue ? { glue } : {}),
			});
		word = '';
		glue = undefined;
	};
	let offset = 0;
	for (const ch of graphemesOf(text)) {
		if (ch === '\n' || ch === '\t' || ch === ' ') {
			flushWord();
			previous = '';
			if (ch === '\n') tokens.push({ kind: 'lineBreak', runIndex, sourceStart: offset });
			else if (ch === '\t') tokens.push({ kind: 'tab', runIndex, sourceStart: offset });
			else tokens.push({ kind: 'space', text: ch, runIndex, sourceStart: offset });
		} else {
			if (word && breakAllowedBetween(previous, ch)) {
				const kept = kinsokuForbids(previous, ch, kinsoku, kinsoku);
				flushWord();
				if (kept) glue = 'kinsoku';
			}
			if (!word) wordStart = offset;
			word += ch;
			previous = ch;
		}
		offset += ch.length;
	}
	flushWord();
	return tokens;
}

/** Appends a forced page or column break marker after a run's tokens. */
export function appendBreakMarker(
	tokens: BreakToken[],
	runIndex: number,
	sourceStart: number,
	kind: 'page' | 'column',
): void {
	tokens.push({ kind: kind === 'page' ? 'pageBreak' : 'columnBreak', runIndex, sourceStart });
}

/**
 * How many of a line's last placed tokens must move to the next line together with
 * `tokens[index]`, which does not fit: the chain of glued tokens ending at it. 0 when the token
 * is not glued, or when the chain reaches back to the line's first token (nothing could stay on
 * the line, so the break falls before `tokens[index]` after all).
 */
export function gluedCarry(
	tokens: readonly BreakToken[],
	index: number,
	placed: readonly BreakToken[],
): number {
	let carry = 0;
	for (let current = index; ; current--) {
		const token = tokens[current];
		if (!token?.glue) return carry;
		const previous = placed[placed.length - 1 - carry];
		if (!previous || previous !== tokens[current - 1]) return carry;
		carry++;
		if (carry >= placed.length) return 0;
	}
}
