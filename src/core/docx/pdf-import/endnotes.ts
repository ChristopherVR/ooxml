/**
 * Endnotes collected at the end of the document: after the last body text, under a short
 * separator rule (Word's) or a "Notes" heading, a run of notes that each start with a mark, every
 * mark answering a raised mark earlier in the text, in order. The notes leave the body and become
 * Word endnotes; the body marks become the endnote references.
 *
 * The rule is kept strict so footnotes are not taken for endnotes: the block runs to the end of
 * the document, every note in it has a body mark, and either a heading introduces it, it does not
 * sit at the foot of its page, or its marks reach back more than a page. Endnotes collected at the
 * end of each section are not recognised (they stay body text).
 */
import type { PdfPageRule } from '../../pdf/index';
import type { Fragment, Line } from './lines';
import {
	groupNotes,
	isFreeMark,
	isSeparatorRule,
	leadingMark,
	noteLines,
	type NoteContent,
} from './note-marks';

export interface EndnotePage {
	lines: Line[];
	rules: readonly PdfPageRule[];
}

/** The numbering Word shows for the endnotes (`w:endnotePr/w:numFmt`). */
export type EndnoteFormat = 'decimal' | 'lowerRoman' | 'upperRoman' | 'lowerLetter';

export interface Endnotes {
	notes: NoteContent[];
	format: EndnoteFormat;
	/** Whether the endnotes were collected at the end of the document or at the end of each section. */
	position?: 'docEnd' | 'sectEnd';
}

const HEADING = /^(end ?notes?|notes?|notes de fin|endnoten|anmerkungen|notas( finales)?|note)$/i;

function roman(value: number): string {
	const digits: [number, string][] = [
		[100, 'c'],
		[90, 'xc'],
		[50, 'l'],
		[40, 'xl'],
		[10, 'x'],
		[9, 'ix'],
		[5, 'v'],
		[4, 'iv'],
		[1, 'i'],
	];
	let out = '';
	for (const [size, text] of digits)
		while (value >= size) {
			out += text;
			value -= size;
		}
	return out;
}

const FORMATS: [EndnoteFormat, (id: number) => string][] = [
	['decimal', String],
	['lowerRoman', roman],
	['upperRoman', (id) => roman(id).toUpperCase()],
	['lowerLetter', (id) => String.fromCharCode(96 + id)],
];

/** Where an endnote block may start: under a separator rule, or after a notes heading. */
interface Start {
	page: number;
	/** The first line of the block on `page` (by baseline order). */
	index: number;
	heading?: Line;
}

function starts(sorted: Line[][], pages: EndnotePage[]): Start[] {
	const found: Start[] = [];
	sorted.forEach((lines, page) => {
		lines.forEach((line, index) => {
			const above = lines[index - 1];
			const ruled = pages[page]!.rules.some(
				(rule) =>
					isSeparatorRule(rule) &&
					rule.y0 < line.top &&
					rule.y0 > (above?.bottom ?? -Infinity) &&
					!!leadingMark(line),
			);
			if (ruled) found.push({ page, index });
			else if (HEADING.test(line.text.trim()) && lines[index + 1] && leadingMark(lines[index + 1]!))
				found.push({ page, index: index + 1, heading: line });
		});
	});
	return found;
}

/**
 * Takes the endnotes off `pages` (in place) when the document ends with them: note lines (and a
 * notes heading) leave `lines`, body marks get their `note`.
 */
export function extractEndnotes(pages: EndnotePage[], bodySize: number): Endnotes | undefined {
	const sorted = pages.map((page) => [...page.lines].sort((a, b) => a.baseline - b.baseline));
	const foot = Math.max(...sorted.flat().map((line) => line.bottom), 0);
	for (const start of starts(sorted, pages)) {
		const block = [
			...sorted[start.page]!.slice(start.index),
			...sorted.slice(start.page + 1).flat(),
		];
		// Endnote text is smaller than the body's, unless a heading says what follows; even then it
		// is no larger than the first note's.
		const limit = start.heading ? block[0]!.size + 0.5 : bodySize * 0.95;
		if (block.some((line) => line.size >= limit)) continue;
		const groups = groupNotes(block);
		if (!groups.length) continue;
		// The text before the block, in reading order, and its free raised marks.
		const before = [
			...sorted.slice(0, start.page).flat(),
			...sorted[start.page]!.slice(0, start.index),
		];
		// A raised mark that starts a line is a note's own mark (a footnote's), not a reference.
		const marks = before.flatMap((line) =>
			line.fragments
				.filter((fragment, index) => index > 0 || !fragment.script)
				.map((fragment) => ({ fragment, page: line.page })),
		);
		let at = 0;
		const references: { fragment: Fragment; page: number }[] = [];
		for (const group of groups) {
			while (at < marks.length && !isFreeMark(marks[at]!.fragment, group.mark)) at++;
			if (at >= marks.length) break;
			references.push(marks[at++]!);
		}
		if (references.length !== groups.length) continue;
		const blockPage = block[0]!.page;
		const atFoot = block[block.length - 1]!.bottom > foot - 1.5 * bodySize;
		const farBack = references.some((reference) => reference.page < blockPage - 1);
		if (!start.heading && atFoot && !farBack) continue;
		const format =
			FORMATS.find(([, mark]) =>
				groups.every((group, index) => group.mark === mark(index + 1)),
			)?.[0] ?? 'decimal';
		const numbered = FORMATS.find(([name]) => name === format)![1];
		const notes = groups.map((group, index): NoteContent => {
			const id = index + 1;
			const custom = group.mark !== numbered(id);
			references[index]!.fragment.note = { id, mark: group.mark, custom, kind: 'endnote' };
			return { id, ...(custom ? { customMark: group.mark } : {}), lines: noteLines(group) };
		});
		const taken = new Set<Line>([...block, ...(start.heading ? [start.heading] : [])]);
		for (const page of pages) page.lines = page.lines.filter((line) => !taken.has(line));
		return { notes, format, position: 'docEnd' };
	}
	return extractSectionEndnotes(sorted, pages, bodySize, foot);
}

/**
 * Endnotes collected at the end of each section: one or more separator- or heading-introduced
 * note blocks whose marks match raised marks in the preceding section.
 */
function extractSectionEndnotes(
	sorted: Line[][],
	pages: EndnotePage[],
	bodySize: number,
	foot: number,
): Endnotes | undefined {
	const candidates = starts(sorted, pages);
	if (!candidates.length) return undefined;
	const allGroups: { group: ReturnType<typeof groupNotes>[number]; ref: Fragment }[] = [];
	const taken = new Set<Line>();
	let prevEndPage = 0;
	let prevEndIndex = 0;
	for (let ci = 0; ci < candidates.length; ci++) {
		const start = candidates[ci]!;
		if (start.page < prevEndPage || (start.page === prevEndPage && start.index < prevEndIndex))
			continue;
		const nextStart = candidates[ci + 1];
		const firstLine = sorted[start.page]?.[start.index];
		if (!firstLine) continue;
		const limit = start.heading ? firstLine.size + 0.5 : bodySize * 0.95;
		const block: Line[] = [];
		let endPage = start.page;
		let endIndex = start.index;
		for (let p = start.page; p < sorted.length; p++) {
			const pageLines = sorted[p]!;
			const from = p === start.page ? start.index : 0;
			const stopAt = nextStart && nextStart.page === p ? nextStart.index : pageLines.length;
			let broke = false;
			for (let idx = from; idx < stopAt; idx++) {
				const line = pageLines[idx]!;
				if (line.size >= limit) {
					broke = true;
					break;
				}
				// On a later page, stop if the text jumps to a non-note paragraph (new section).
				if (p > start.page && idx === 0 && !leadingMark(line) && line.size >= bodySize * 0.95) {
					broke = true;
					break;
				}
				block.push(line);
				endPage = p;
				endIndex = idx + 1;
			}
			// Section endnotes end at a section break (typically a page boundary).
			if (broke || p === start.page) break;
		}
		if (!block.length) continue;
		const groups = groupNotes(block);
		if (!groups.length) continue;
		// Body lines in this section, before the note block.
		const before: Line[] = [];
		for (let p = prevEndPage; p <= start.page; p++) {
			const pageLines = sorted[p]!;
			const from = p === prevEndPage ? prevEndIndex : 0;
			const to = p === start.page ? start.index : pageLines.length;
			for (let idx = from; idx < to; idx++) before.push(pageLines[idx]!);
		}
		const marks = before.flatMap((line) =>
			line.fragments
				.filter((fragment, index) => index > 0 || !fragment.script)
				.map((fragment) => ({ fragment, page: line.page })),
		);
		let at = 0;
		const references: { fragment: Fragment; page: number }[] = [];
		for (const group of groups) {
			while (at < marks.length && !isFreeMark(marks[at]!.fragment, group.mark)) at++;
			if (at >= marks.length) break;
			references.push(marks[at++]!);
		}
		if (references.length !== groups.length) continue;
		const blockPage = block[0]!.page;
		const atFoot = block[block.length - 1]!.bottom > foot - 1.5 * bodySize;
		const farBack = references.some((reference) => reference.page < blockPage - 1);
		if (!start.heading && atFoot && !farBack) continue;
		groups.forEach((group, index) => allGroups.push({ group, ref: references[index]!.fragment }));
		for (const line of block) taken.add(line);
		if (start.heading) taken.add(start.heading);
		prevEndPage = endPage;
		prevEndIndex = endIndex;
	}
	if (!allGroups.length) return undefined;
	const format =
		FORMATS.find(([, mark]) =>
			allGroups.every(({ group }, index) => group.mark === mark(index + 1)),
		)?.[0] ?? 'decimal';
	const numbered = FORMATS.find(([name]) => name === format)![1];
	const notes = allGroups.map(({ group, ref }, index): NoteContent => {
		const id = index + 1;
		const custom = group.mark !== numbered(id);
		ref.note = { id, mark: group.mark, custom, kind: 'endnote' };
		return { id, ...(custom ? { customMark: group.mark } : {}), lines: noteLines(group) };
	});
	for (const page of pages) page.lines = page.lines.filter((line) => !taken.has(line));
	return { notes, format, position: 'sectEnd' };
}
