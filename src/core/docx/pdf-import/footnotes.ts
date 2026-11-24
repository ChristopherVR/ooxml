/**
 * Footnotes: at the bottom of a page, a run of lines smaller than the body text, each note
 * starting with a mark (a raised number or symbol, or a number set before the text) that a raised
 * mark in the page's body text matches, often under a short separator rule. Matched notes leave
 * the body and become Word footnotes; the body mark becomes the footnote reference.
 */
import type { PdfPageRule } from '../../pdf/index';
import type { Fragment, Line } from './lines';

export interface FootnoteContent {
	id: number;
	/** The mark, when it is not the note's number (a symbol, or numbering that restarts). */
	customMark?: string;
	/** The note's lines, its mark removed. */
	lines: Line[];
}

export interface FootnotePage {
	lines: Line[];
	height: number;
	rules: readonly PdfPageRule[];
	/** The lowest point of the notes taken from the page (they still take room in the layout). */
	notesBottom?: number;
}

const MARK = /^(\d{1,3}|[*\u2020\u2021\u00a7\u00b6]{1,3}|[a-z])$/;

/** The mark a note line starts with, and the fragments after it, if it starts with one. */
function leadingMark(line: Line): { mark: string; rest: Fragment[]; raised: boolean } | undefined {
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

/** The body fragment carrying reference mark `mark` (raised), searching from the bottom up. */
function referenceFor(lines: Line[], mark: string): Fragment | undefined {
	for (let index = lines.length - 1; index >= 0; index--) {
		const found = [...lines[index]!.fragments]
			.reverse()
			.find(
				(fragment) =>
					fragment.script === 'super' && fragment.text.trim() === mark && !fragment.note,
			);
		if (found) return found;
	}
	return undefined;
}

/**
 * Takes the footnotes off `pages` (in place): note lines leave `lines`, body marks get their
 * `note`. Returns the notes in document order.
 */
export function extractFootnotes(pages: FootnotePage[], bodySize: number): FootnoteContent[] {
	const notes: FootnoteContent[] = [];
	for (const page of pages) {
		const sorted = [...page.lines].sort((a, b) => a.baseline - b.baseline);
		let start = sorted.length;
		while (
			start > 0 &&
			sorted[start - 1]!.size < bodySize * 0.92 &&
			sorted[start - 1]!.top > page.height * 0.45
		)
			start--;
		const zone = sorted.slice(start);
		const firstNote = zone.findIndex((line) => leadingMark(line));
		if (firstNote < 0) continue;
		const noteLines = zone.slice(firstNote);
		const body = sorted.slice(0, start + firstNote);
		const zoneTop = noteLines[0]!.top;
		const separator = page.rules.some(
			(rule) =>
				rule.y1 - rule.y0 < 2 &&
				rule.x1 - rule.x0 < 250 &&
				rule.y0 < zoneTop &&
				rule.y0 > zoneTop - 3 * bodySize,
		);
		const groups: { mark: string; lines: Line[]; first: Fragment[]; raised: boolean }[] = [];
		for (const line of noteLines) {
			const lead = leadingMark(line);
			if (lead)
				groups.push({ mark: lead.mark, lines: [line], first: lead.rest, raised: lead.raised });
			else groups[groups.length - 1]?.lines.push(line);
		}
		const taken: Line[] = [];
		for (const group of groups) {
			const reference = referenceFor(body, group.mark);
			// A number on the baseline is a note mark only under a separator rule.
			if (!reference || (!group.raised && !separator)) continue;
			const id = notes.length + 1;
			reference.note = { id, mark: group.mark, custom: group.mark !== String(id) };
			const [head, ...tail] = group.lines;
			const first = group.first.map((fragment) => ({ ...fragment }));
			if (first[0]) first[0].text = first[0].text.replace(/^\s+/, '');
			const left = first[0]?.x ?? head!.left;
			notes.push({
				id,
				...(reference.note.custom ? { customMark: group.mark } : {}),
				lines: [
					{
						...head!,
						fragments: first,
						text: first.map((fragment) => fragment.text).join(''),
						left,
					},
					...tail,
				],
			});
			taken.push(...group.lines);
		}
		if (!taken.length) continue;
		page.notesBottom = Math.max(...taken.map((line) => line.bottom));
		page.lines = page.lines.filter((line) => !taken.includes(line));
	}
	return notes;
}
