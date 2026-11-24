/**
 * What footnotes and endnotes share: the mark a note starts with, the separator rule above a note
 * area, the body mark a note answers, and the note's lines once its mark is taken off.
 */
import type { PdfPageRule } from '../../pdf/index';
import type { Fragment, Line } from './lines';
import { markerOf } from './list-markers';
import { buildParagraphs, type ParagraphBlock, type TextArea } from './paragraphs';

export interface NoteContent {
	id: number;
	/** The mark, when it is not the note's number (a symbol, or numbering that restarts). */
	customMark?: string;
	/** The note's lines, its mark removed (from more than one page when the note continues). */
	lines: Line[];
}

/** A note in a note area: its mark and lines, before it is matched to a body mark. */
export interface NoteGroup {
	mark: string;
	lines: Line[];
	/** The first line's fragments after the mark. */
	first: Fragment[];
	raised: boolean;
}

const MARK = /^(\d{1,3}|[ivxlc]{1,6}|[IVXLC]{1,6}|[*\u2020\u2021\u00a7\u00b6]{1,3}|[a-z])$/;

/** The mark a note line starts with, and the fragments after it, if it starts with one. */
export function leadingMark(
	line: Line,
): { mark: string; rest: Fragment[]; raised: boolean } | undefined {
	const [first, ...others] = line.fragments.filter((fragment) => fragment.text.length);
	if (!first) return undefined;
	const text = first.text.trim();
	if (first.script === 'super' && MARK.test(text))
		return { mark: text, rest: others, raised: true };
	// A number set on the baseline before the note text: "1 Text" or "1. Text".
	const match = /^\s*(\d{1,3})[.)]?\s+(?=\S)/.exec(first.text);
	if (match && !first.script) {
		const remainder = { ...first, text: first.text.slice(match[0].length) };
		return { mark: match[1]!, rest: [remainder, ...others], raised: false };
	}
	return undefined;
}

/** Whether `fragment` is a raised body mark `mark` that no note has taken yet. */
export const isFreeMark = (fragment: Fragment, mark: string) =>
	fragment.script === 'super' && fragment.text.trim() === mark && !fragment.note;

/** The body fragment carrying reference mark `mark` (raised), searching from the bottom up. */
export function referenceFor(lines: readonly Line[], mark: string): Fragment | undefined {
	for (let index = lines.length - 1; index >= 0; index--) {
		const found = [...lines[index]!.fragments].reverse().find((item) => isFreeMark(item, mark));
		if (found) return found;
	}
	return undefined;
}

/** A thin horizontal rule (a note separator), shorter than `maxWidth`. */
export const isSeparatorRule = (rule: PdfPageRule, maxWidth = 250) =>
	rule.y1 - rule.y0 < 2 && rule.x1 - rule.x0 > 20 && rule.x1 - rule.x0 < maxWidth;

/** Whether a separator rule lies just above `top` (within three body lines). */
export const separatorAbove = (rules: readonly PdfPageRule[], top: number, bodySize: number) =>
	rules.some((rule) => isSeparatorRule(rule) && rule.y0 < top && rule.y0 > top - 3 * bodySize);

/** Splits note lines into notes, each starting at a marked line; unmarked lines go on the one before. */
export function groupNotes(lines: readonly Line[]): NoteGroup[] {
	const groups: NoteGroup[] = [];
	for (const line of lines) {
		const lead = leadingMark(line);
		if (lead)
			groups.push({ mark: lead.mark, lines: [line], first: lead.rest, raised: lead.raised });
		else groups[groups.length - 1]?.lines.push(line);
	}
	return groups;
}

/** The note's lines with its mark taken off the first one. */
export function noteLines(group: NoteGroup): Line[] {
	const [head, ...tail] = group.lines;
	const first = group.first.map((fragment) => ({ ...fragment }));
	if (first[0]) first[0].text = first[0].text.replace(/^\s+/, '');
	return [
		{
			...head!,
			fragments: first,
			text: first.map((fragment) => fragment.text).join(''),
			left: first[0]?.x ?? head!.left,
		},
		...tail,
	];
}

const TERMINAL = /[.!?:;\u2026\u00bb\u201d"')\]]\s*$/;

/**
 * A note's paragraphs. A note that a page break cut is built page by page; its paragraph goes on
 * across the break unless the text before it ends a sentence.
 */
export function noteParagraphs(
	lines: readonly Line[],
	area: TextArea,
	bodySize: number,
): ParagraphBlock[] {
	const pieces: Line[][] = [];
	for (const line of lines) {
		const piece = pieces[pieces.length - 1];
		if (piece && piece[0]!.page === line.page) piece.push(line);
		else pieces.push([line]);
	}
	const blocks: ParagraphBlock[] = [];
	for (const piece of pieces) {
		const built = buildParagraphs(piece, area, bodySize).map(({ heading: _heading, ...block }) => ({
			...block,
			indentLeft: 0,
			firstLine: 0,
			spaceBefore: 0,
		}));
		const last = blocks[blocks.length - 1];
		const next = built[0];
		if (
			last &&
			next &&
			!TERMINAL.test(last.lines[last.lines.length - 1]!.text) &&
			!markerOf(next.lines[0]!)
		) {
			const tail = last.fragments[last.fragments.length - 1];
			if (tail) {
				const hyphen =
					/[-\u00ad]$/.test(tail.text) && /^\s*[a-z\u00df-\u00ff]/.test(next.lines[0]!.text);
				tail.text = hyphen ? tail.text.replace(/[-\u00ad]$/, '') : `${tail.text} `;
			}
			last.fragments.push(...next.fragments);
			last.lines.push(...next.lines);
			built.shift();
		}
		blocks.push(...built);
	}
	return blocks;
}
