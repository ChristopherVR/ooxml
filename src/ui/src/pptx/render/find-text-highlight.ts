/**
 * Which characters of a slide's text the Find bar highlights, shared by every binding's text
 * renderer: the matches of the active slide grouped by element and text segment, a segment's text
 * split into plain and highlighted chunks, and the runs of `buildParagraphs` split at match
 * boundaries. DOM-free; a binding only wraps a highlighted piece in
 * `<mark data-pptx-find-match="match|current">` styled with {@link findMatchMarkStyle}.
 *
 * Only a top-level shape's own text is highlighted (as React always did): matches in a table
 * cell, a SmartArt node, a group member, a chart or the speaker notes select their element but
 * draw no mark.
 *
 * @module render/find-text-highlight
 */
import type { FindResult } from './find-replace';
import type { ParagraphRun, RenderParagraph } from './paragraph-types';
import type { RunStyle } from './text-run-style';
import type { ScriptFontPiece } from './text-script-fonts';

/** The attribute on a highlighted piece of text: `current` for the current match, else `match`. */
export const FIND_MATCH_ATTRIBUTE = 'data-pptx-find-match';

/** A highlighted piece: one of the matches, or the current match. */
export type FindMatchKind = 'match' | 'current';

/** One match inside a text segment, in characters of the segment's text. */
export interface FindTextHighlight {
	startOffset: number;
	length: number;
	/** The match the Find bar is on. */
	isCurrent: boolean;
}

/** The highlights of one element, by `textSegments` index. */
export type ElementFindHighlights = ReadonlyMap<number, readonly FindTextHighlight[]>;

/** The highlights of a slide, by element id. */
export type SlideFindHighlights = ReadonlyMap<string, ElementFindHighlights>;

/** A piece of text and whether it is highlighted. */
export interface FindTextChunk {
	text: string;
	match?: FindMatchKind;
}

/** Whether a match addresses its element's own `textSegments` (not a cell, node, member, ...). */
function inShapeText(match: FindResult): boolean {
	return (
		!match.notes &&
		!match.tableCell &&
		match.smartArtNodeId === undefined &&
		match.elementPath === undefined &&
		!match.chartTitle &&
		match.chartSeriesIndex === undefined &&
		match.chartCategoryIndex === undefined
	);
}

/** The matches of slide `slideIndex`, grouped by element and segment, the current one marked. */
export function findTextHighlightsByElement(
	matches: readonly FindResult[],
	currentIndex: number,
	slideIndex: number,
): Map<string, Map<number, FindTextHighlight[]>> {
	const byElement = new Map<string, Map<number, FindTextHighlight[]>>();
	matches.forEach((match, index) => {
		if (match.slideIndex !== slideIndex || !inShapeText(match)) {
			return;
		}
		let bySegment = byElement.get(match.elementId);
		if (!bySegment) {
			bySegment = new Map();
			byElement.set(match.elementId, bySegment);
		}
		let list = bySegment.get(match.segmentIndex);
		if (!list) {
			list = [];
			bySegment.set(match.segmentIndex, list);
		}
		list.push({
			startOffset: match.startOffset,
			length: match.length,
			isCurrent: index === currentIndex,
		});
	});
	return byElement;
}

/**
 * Split `text` into plain and highlighted chunks. Highlights are clipped to the text, and an
 * overlapping match (the search counts them) only highlights what the previous one did not.
 */
export function splitTextByFindHighlights(
	text: string,
	highlights: readonly FindTextHighlight[] | undefined,
): FindTextChunk[] {
	const sorted = [...(highlights ?? [])].sort((a, b) => a.startOffset - b.startOffset);
	const chunks: FindTextChunk[] = [];
	let cursor = 0;
	for (const highlight of sorted) {
		const start = Math.max(highlight.startOffset, cursor);
		const end = Math.min(highlight.startOffset + highlight.length, text.length);
		if (end <= start) {
			continue;
		}
		if (start > cursor) {
			chunks.push({ text: text.slice(cursor, start) });
		}
		chunks.push({ text: text.slice(start, end), match: highlight.isCurrent ? 'current' : 'match' });
		cursor = end;
	}
	if (cursor < text.length || chunks.length === 0) {
		chunks.push({ text: text.slice(cursor) });
	}
	return chunks;
}

/** The inline style of a highlighted piece (the current match is orange, the others yellow). */
export function findMatchMarkStyle(kind: FindMatchKind): RunStyle {
	return {
		backgroundColor: kind === 'current' ? '#f97316' : '#facc15',
		color: 'inherit',
		borderRadius: '2px',
	};
}

/** The pieces of `pieces` covering characters `[from, to)` of their joined text. */
function slicePieces(
	pieces: readonly ScriptFontPiece[],
	from: number,
	to: number,
): ScriptFontPiece[] {
	const out: ScriptFontPiece[] = [];
	let offset = 0;
	for (const piece of pieces) {
		const start = Math.max(from, offset);
		const end = Math.min(to, offset + piece.text.length);
		if (end > start) {
			out.push({ ...piece, text: piece.text.slice(start - offset, end - offset) });
		}
		offset += piece.text.length;
	}
	return out;
}

/** One run split at the match boundaries that fall inside it. */
function highlightRun(run: ParagraphRun, highlights: ElementFindHighlights): ParagraphRun[] {
	if (run.equation || run.segmentIndex === undefined || !run.text) {
		return [run];
	}
	const segment = highlights.get(run.segmentIndex);
	if (!segment?.length) {
		return [run];
	}
	const base = run.charStart ?? 0;
	const local = segment.map((highlight) => ({
		...highlight,
		startOffset: highlight.startOffset - base,
	}));
	const chunks = splitTextByFindHighlights(run.text, local);
	const marked = chunks.filter((chunk) => chunk.match);
	if (marked.length === 0) {
		return [run];
	}
	// A run whose text is laid out as a whole (a ruby base, measured tab stops) is not split:
	// it is highlighted entirely when a match touches it.
	if (run.ruby || run.tabLines) {
		const current = marked.some((chunk) => chunk.match === 'current');
		return [{ ...run, findMatch: current ? 'current' : 'match' }];
	}
	const out: ParagraphRun[] = [];
	let offset = 0;
	for (const chunk of chunks) {
		const end = offset + chunk.text.length;
		const piece: ParagraphRun = { ...run, text: chunk.text, charStart: base + offset };
		if (run.scriptRuns) {
			piece.scriptRuns = slicePieces(run.scriptRuns, offset, end);
		}
		if (run.underlineWordPieces) {
			piece.underlineWordPieces = slicePieces(run.underlineWordPieces, offset, end);
		}
		if (chunk.match) {
			piece.findMatch = chunk.match;
		}
		out.push(piece);
		offset = end;
	}
	return out;
}

/**
 * The paragraphs of `buildParagraphs` with each run that holds a match split at the match
 * boundaries, the highlighted pieces carrying `findMatch`. Returns `paragraphs` itself when the
 * element has no highlights.
 */
export function highlightFindMatches(
	paragraphs: RenderParagraph[],
	highlights: ElementFindHighlights | undefined,
): RenderParagraph[] {
	if (!highlights || highlights.size === 0) {
		return paragraphs;
	}
	return paragraphs.map((paragraph) => ({
		...paragraph,
		runs: paragraph.runs.flatMap((run) => highlightRun(run, highlights)),
	}));
}
