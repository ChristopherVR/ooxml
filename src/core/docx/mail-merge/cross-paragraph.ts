// Merge fields whose code or result crosses paragraph marks, such as an IF whose branch text holds
// several paragraphs. The paragraphs a complex field spans are joined into one run sequence with
// a private-use character for each paragraph mark, evaluated like one paragraph, then split
// again. Each split paragraph takes the properties of the paragraph whose mark ends it, as a
// Word paragraph takes its formatting from its paragraph mark.
import { cellSourceBlocks, mapStoryBlocks } from '../cell-blocks';
import type { Block, Paragraph, TableCell, TextRun } from '../model';

const MARK = 0xe100;
/** At most this many paragraphs are joined; a longer field keeps its source result. */
export const MAX_JOINED_PARAGRAPHS = 256;
const MARKS = /[-]/u;
const SPLIT = /([-])/u;

function openAfter(runs: readonly TextRun[], open: number): number {
	for (const run of runs) {
		if (run.fieldChar === 'begin') open++;
		else if (run.fieldChar === 'end' && open > 0) open--;
	}
	return open;
}

function blockOpenAfter(block: Block, open: number): number {
	if (block.type === 'paragraph') return openAfter(block.runs, open);
	for (const row of block.rows)
		for (const cell of row)
			for (const item of cellSourceBlocks(cell)) open = blockOpenAfter(item, open);
	return open;
}

const usesMarks = (paragraph: Paragraph) =>
	paragraph.runs.some((run) => MARKS.test(run.text) || MARKS.test(run.fieldCode ?? ''));

/** One paragraph holding the runs of `group`, with a mark run at each paragraph boundary. */
export function joinParagraphs(group: readonly Paragraph[]): Paragraph {
	const runs: TextRun[] = [];
	// Whether each open field is still in its code (no `separate` yet).
	const fields: boolean[] = [];
	group.forEach((paragraph, index) => {
		for (const run of paragraph.runs) {
			runs.push(run);
			if (run.fieldChar === 'begin') fields.push(true);
			else if (run.fieldChar === 'separate' && fields.length) fields[fields.length - 1] = false;
			else if (run.fieldChar === 'end') fields.pop();
		}
		if (index === group.length - 1) return;
		const mark = String.fromCharCode(MARK + index);
		runs.push(fields.at(-1) ? { text: '', fieldCode: mark } : { text: mark });
	});
	return { ...group[0]!, runs };
}

/** Splits a joined paragraph at its remaining marks; `group` supplies each paragraph's properties. */
export function splitParagraphs(
	joined: Paragraph,
	group: readonly Paragraph[],
	used = new Set<string>(),
): Paragraph[] {
	const output: Paragraph[] = [];
	let runs: TextRun[] = [];
	const close = (owner: number) => {
		const source = group[owner] ?? group.at(-1)!;
		let id = source.id;
		for (let copy = 1; used.has(id); copy++) id = `${source.id}-merge-${copy}`;
		used.add(id);
		output.push({ ...source, id, runs: runs.length ? runs : [{ text: '' }] });
		runs = [];
	};
	for (const run of joined.runs) {
		const code = run.fieldCode;
		if (code !== undefined && code.length === 1 && MARKS.test(code) && !run.text) {
			close(code.charCodeAt(0) - MARK);
			continue;
		}
		if (!MARKS.test(run.text)) {
			runs.push(run);
			continue;
		}
		for (const part of run.text.split(SPLIT)) {
			if (part.length === 1 && MARKS.test(part)) close(part.charCodeAt(0) - MARK);
			else if (part) runs.push({ ...run, text: part });
		}
	}
	close(group.length - 1);
	return output;
}

function mapCell(cell: TableCell, paragraph: (value: Paragraph) => Paragraph): TableCell {
	if (cell.sourceBlocks) {
		const original = cellSourceBlocks(cell);
		const source = mapStoryBlocks(original, paragraph);
		if (source === original) return cell;
		return {
			...cell,
			paragraphs: source.filter((item) => item.type === 'paragraph'),
			sourceBlocks: source,
		};
	}
	const paragraphs = mapMergeStory(cell.paragraphs, paragraph);
	return paragraphs === cell.paragraphs
		? cell
		: { ...cell, paragraphs: paragraphs.filter((item) => item.type === 'paragraph') };
}

/**
 * Maps a story's paragraphs in order like `mapStoryBlocks`, passing each run of paragraphs that a
 * complex field spans to `paragraph` as one joined paragraph. Fields that cross a table, are left
 * open or span too many paragraphs reach `paragraph` one paragraph at a time.
 */
export function mapMergeStory(
	blocks: readonly Block[],
	paragraph: (value: Paragraph) => Paragraph,
): Block[] {
	const output: Block[] = [];
	let changed = false;
	let open = 0;
	const used = new Set<string>();
	for (let index = 0; index < blocks.length; index++) {
		const block = blocks[index]!;
		if (block.type === 'table') {
			const rows = block.rows.map((row) => row.map((cell) => mapCell(cell, paragraph)));
			const same = rows.every((row, r) => row.every((cell, c) => cell === block.rows[r]![c]));
			output.push(same ? block : { ...block, rows });
			changed ||= !same;
			open = blockOpenAfter(block, open);
			continue;
		}
		let end = index;
		if (open === 0 && openAfter(block.runs, 0) > 0) {
			let depth = openAfter(block.runs, 0);
			while (depth > 0 && end + 1 < blocks.length && end - index + 1 < MAX_JOINED_PARAGRAPHS) {
				const next = blocks[end + 1]!;
				if (next.type !== 'paragraph') break;
				end++;
				depth = openAfter(next.runs, depth);
			}
			if (depth > 0) end = index;
		}
		const group = blocks.slice(index, end + 1) as Paragraph[];
		if (group.length > 1 && !group.some(usesMarks)) {
			const joined = joinParagraphs(group);
			const projected = paragraph(joined);
			if (projected === joined) output.push(...group);
			else {
				output.push(...splitParagraphs(projected, group, used));
				changed = true;
			}
			index = end;
			continue;
		}
		const next = paragraph(block);
		changed ||= next !== block;
		output.push(next);
		used.add(next.id);
		open = openAfter(block.runs, open);
	}
	return changed ? output : (blocks as Block[]);
}

/** A field code with the paragraph marks outside its quoted text read as spaces. */
export function markSpaces(instruction: string): string {
	if (!MARKS.test(instruction)) return instruction;
	let quoted = false;
	let output = '';
	for (let at = 0; at < instruction.length; at++) {
		const char = instruction[at]!;
		if (char === '\\' && at + 1 < instruction.length) {
			output += char + instruction[++at];
			continue;
		}
		if (char === '"') quoted = !quoted;
		output += !quoted && MARKS.test(char) ? ' ' : char;
	}
	return output;
}
