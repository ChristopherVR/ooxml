/**
 * Word wrapping for the WordArt glyph envelope layout
 * (`text-warp-envelope-block.ts`).
 *
 * PowerPoint wraps a warped text body at the text box width first (unless
 * `a:bodyPr/@wrap="none"`) and only then maps the wrapped block onto the
 * preset's envelope, so a long paragraph in an ordinary text box becomes
 * several rows that are stretched together, not one squeezed row.
 *
 * Rows break only at whitespace. A word wider than the available width stays
 * whole on its own row (it is squeezed by the envelope instead), and a hard
 * line break (`a:br`, carried as a `'\n'` run) always starts a new row.
 * Widths are the nominal (unwarped) glyph advances the layout already uses.
 */
import { DEFAULT_BODY_INSET_LR_PX } from '../constants';
import type { EnvelopeSegmentInput } from './text-warp-envelope-types';

/** The `a:bodyPr` properties that decide whether, and how wide, the block wraps. */
export interface EnvelopeTextBody {
	/** Left inset in px (`a:bodyPr/@lIns`); the OOXML default when absent. */
	bodyInsetLeft?: number;
	/** Right inset in px (`a:bodyPr/@rIns`); the OOXML default when absent. */
	bodyInsetRight?: number;
	/** `a:bodyPr/@wrap`: `'none'` keeps each paragraph on one row. */
	textWrap?: string;
}

/** One character of a paragraph with the run it came from and its advance. */
export interface EnvelopeWrapItem {
	seg: EnvelopeSegmentInput;
	char: string;
	advance: number;
}

/**
 * Whitespace a row may break at. No-break spaces (U+00A0, U+2007, U+202F)
 * hold their words together, as in ordinary text.
 */
const WHITESPACE_RE = /^(?![\u00A0\u2007\u202F])\s$/u;
const HARD_BREAK = '\n';
/** Slack for summed floating-point advances that exactly fill the width. */
const FIT_EPSILON = 1e-6;

/**
 * The width a `boxWidth`-wide shape wraps its text at: the box less its left
 * and right insets, or `undefined` (no wrapping) when `body` is absent or
 * says `wrap="none"`.
 */
export function envelopeWrapWidth(
	boxWidth: number,
	body: EnvelopeTextBody | undefined,
): number | undefined {
	if (!body || body.textWrap === 'none') {
		return undefined;
	}
	const left = body.bodyInsetLeft ?? DEFAULT_BODY_INSET_LR_PX;
	const right = body.bodyInsetRight ?? DEFAULT_BODY_INSET_LR_PX;
	return Math.max(0, boxWidth - left - right);
}

function sumAdvances(items: readonly EnvelopeWrapItem[]): number {
	return items.reduce((sum, item) => sum + item.advance, 0);
}

/** Greedy word wrap of one hard-break-free piece of a paragraph. */
function wrapPiece(items: EnvelopeWrapItem[], maxWidth: number): EnvelopeWrapItem[][] {
	const rows: EnvelopeWrapItem[][] = [];
	let row: EnvelopeWrapItem[] = [];
	let rowWidth = 0;
	let rowHasWord = false;
	let pending: EnvelopeWrapItem[] = [];
	let i = 0;
	while (i < items.length) {
		if (WHITESPACE_RE.test(items[i].char)) {
			pending.push(items[i]);
			i++;
			continue;
		}
		const word: EnvelopeWrapItem[] = [];
		while (i < items.length && !WHITESPACE_RE.test(items[i].char)) {
			word.push(items[i]);
			i++;
		}
		const wordWidth = sumAdvances(word);
		const pendingWidth = sumAdvances(pending);
		if (rowHasWord && rowWidth + pendingWidth + wordWidth > maxWidth + FIT_EPSILON) {
			// Soft break: the whitespace at the break hangs off the row end and
			// is dropped, so it neither renders nor widens the block.
			rows.push(row);
			row = word;
			rowWidth = wordWidth;
		} else {
			row.push(...pending, ...word);
			rowWidth += pendingWidth + wordWidth;
		}
		rowHasWord = true;
		pending = [];
	}
	row.push(...pending);
	rows.push(row);
	return rows;
}

/**
 * Break one paragraph's characters into rows no wider than `maxWidth`
 * (see the module comment). With `maxWidth` `undefined` the paragraph stays
 * one row, unchanged.
 */
export function wrapEnvelopeParagraph(
	items: EnvelopeWrapItem[],
	maxWidth: number | undefined,
): EnvelopeWrapItem[][] {
	if (maxWidth === undefined) {
		return [items];
	}
	const rows: EnvelopeWrapItem[][] = [];
	let piece: EnvelopeWrapItem[] = [];
	for (const item of items) {
		if (item.char === HARD_BREAK) {
			rows.push(...wrapPiece(piece, maxWidth));
			piece = [];
		} else {
			piece.push(item);
		}
	}
	rows.push(...wrapPiece(piece, maxWidth));
	return rows;
}
