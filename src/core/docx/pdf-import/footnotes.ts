/**
 * Footnotes: at the bottom of a page, a run of lines smaller than the body text, each note
 * starting with a mark (a raised number or symbol, or a number set before the text) that a raised
 * mark in the page's body text matches, often under a short separator rule. Matched notes leave
 * the body and become Word footnotes; the body mark becomes the footnote reference.
 *
 * A note can run over a page break: Word continues it at the top of the next page's note area,
 * under the continuation separator and before that page's own notes, so unmarked note lines
 * under a separator go on with the previous page's last note. A note can also start on the page
 * after its mark when the mark's page has no room left; its mark is then looked for on the
 * previous page.
 */
import type { PdfPageRule } from '../../pdf/index';
import type { Fragment, Line } from './lines';
import {
	groupNotes,
	isSeparatorRule,
	leadingMark,
	noteLines,
	referenceFor,
	separatorAbove,
	type NoteContent,
} from './note-marks';

export type FootnoteContent = NoteContent;

export interface FootnotePage {
	lines: Line[];
	height: number;
	rules: readonly PdfPageRule[];
	/** The lowest point of the notes taken from the page (they still take room in the layout). */
	notesBottom?: number;
}

interface PreviousPage {
	/** The page's body lines, top to bottom. */
	body: Line[];
	/** The note the page's note area ends with, which the next page may continue. */
	last?: FootnoteContent;
	hadNotes: boolean;
}

/** The small lines at the foot of a page that may be notes, top to bottom. */
function noteZone(sorted: Line[], height: number, bodySize: number): number {
	let start = sorted.length;
	while (
		start > 0 &&
		sorted[start - 1]!.size < bodySize * 0.92 &&
		sorted[start - 1]!.top > height * 0.45
	)
		start--;
	return start;
}

/** The trailing lines of `before` that sit under a separator: a note going on from the last page. */
function continuedLines(before: Line[], rules: readonly PdfPageRule[], bodySize: number): Line[] {
	for (let index = before.length - 1; index >= 0; index--) {
		const line = before[index]!;
		const above = before[index - 1]?.bottom ?? -Infinity;
		// The continuation separator may run across the whole column.
		const rule = rules.some(
			(item) =>
				isSeparatorRule(item, Infinity) &&
				item.y0 < line.top &&
				item.y0 > Math.max(above, line.top - 3 * bodySize),
		);
		if (rule) return before.slice(index);
	}
	return [];
}

/** The mark on the previous page that a note starting on this page answers, if any. */
function deferredReference(previous: PreviousPage | undefined, mark: string): Fragment | undefined {
	if (!previous) return undefined;
	const reference = referenceFor(previous.body, mark);
	if (!reference) return undefined;
	// The mark's page had a full note area, or the mark sits on one of its last lines.
	const line = previous.body.findIndex((item) => item.fragments.includes(reference));
	return previous.hadNotes || line >= previous.body.length - 3 ? reference : undefined;
}

/**
 * Takes the footnotes off `pages` (in place): note lines leave `lines`, body marks get their
 * `note`. Returns the notes in document order.
 */
export function extractFootnotes(pages: FootnotePage[], bodySize: number): FootnoteContent[] {
	const notes: FootnoteContent[] = [];
	let previous: PreviousPage | undefined;
	for (const page of pages) {
		const sorted = [...page.lines].sort((a, b) => a.baseline - b.baseline);
		const zone = sorted.slice(noteZone(sorted, page.height, bodySize));
		const firstNote = zone.findIndex((line) => leadingMark(line));
		const marked = firstNote < 0 ? [] : zone.slice(firstNote);
		const before = firstNote < 0 ? zone : zone.slice(0, firstNote);
		const continuing = previous?.last;
		const continued = continuing ? continuedLines(before, page.rules, bodySize) : [];
		const body = sorted.filter((line) => !marked.includes(line) && !continued.includes(line));
		const taken: Line[] = [];
		// The note that unmarked or unmatched note lines go on with.
		let current: FootnoteContent | undefined;
		if (continuing && continued.length) {
			continuing.lines.push(...continued);
			taken.push(...continued);
			current = continuing;
		}
		const separator = marked.length > 0 && separatorAbove(page.rules, marked[0]!.top, bodySize);
		let last: FootnoteContent | undefined = current;
		for (const group of groupNotes(marked)) {
			const reference = referenceFor(body, group.mark) ?? deferredReference(previous, group.mark);
			// A number on the baseline is a note mark only under a separator rule.
			if (!reference || (!group.raised && !separator)) {
				// A line of a note that starts with a number goes on with the note before it.
				if (current && !group.raised) {
					current.lines.push(...group.lines);
					taken.push(...group.lines);
				} else current = undefined;
				last = current;
				continue;
			}
			const id = notes.length + 1;
			reference.note = { id, mark: group.mark, custom: group.mark !== String(id) };
			current = {
				id,
				...(reference.note.custom ? { customMark: group.mark } : {}),
				lines: noteLines(group),
			};
			notes.push(current);
			last = current;
			taken.push(...group.lines);
		}
		if (taken.length) {
			page.notesBottom = Math.max(...taken.map((line) => line.bottom));
			page.lines = page.lines.filter((line) => !taken.includes(line));
		}
		previous = {
			body: body.filter((line) => !taken.includes(line)),
			...(last ? { last } : {}),
			hadNotes: taken.length > 0,
		};
	}
	return notes;
}
