// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The fields of a paragraph's runs, with the runs holding their cached results, and the rewrite
// that gives a field a new result. Every field whose result is computed (STYLEREF, DOCPROPERTY,
// SYMBOL...) reads and writes results through these.
import { isFieldLocked } from './field-lock';
import type { Block, Paragraph, Table, TextRun } from './model';
import type { RunSpan } from './style-ref';

/** A field in a paragraph's runs; `from` and `to` span the whole field. */
export interface FieldResultSpan extends RunSpan {
	instr: string;
	locked: boolean;
	/** The runs holding the result. */
	runs: number[];
	/** The run holding the instruction of a complex field. */
	code?: number;
	/** The instruction run holding the first letter of the field name. */
	name?: number;
	/** Where a result goes when there is none: just after the `separate` marker. */
	insertAt: number;
	/** Its result is text a revision removes: kept as it is until the revision is settled. */
	removed?: true;
	/**
	 * A complex field saved without a `separate` marker, and so without a result: Word saves
	 * AUTONUM, LISTNUM and SYMBOL that way and draws what they show. Only listed on request.
	 */
	resultless?: true;
}

/** Whether a revision takes a run's text out of the document: deleted, or moved away. */
const removed = (run: TextRun) =>
	run.revision?.kind === 'delete' || run.revision?.kind === 'moveFrom';

/**
 * Complex fields closed in this paragraph, and simple fields, with the runs of their results. A
 * result run a revision removes is not part of the result: its text is gone, and a new result
 * goes into the text the revisions keep (a result they remove whole stays as it is). Complex
 * fields without a `separate` marker are left out unless `resultless` asks for them.
 */
export function fieldResultSpans(
	runs: readonly TextRun[],
	options: { resultless?: boolean } = {},
): FieldResultSpan[] {
	const found: FieldResultSpan[] = [];
	const open: {
		begin: number;
		code: string;
		codeRun?: number;
		nameRun?: number;
		separate?: number;
		locked: boolean;
	}[] = [];
	runs.forEach((run, index) => {
		const top = open.at(-1);
		if (run.fieldChar === 'begin')
			open.push({ begin: index, code: '', locked: isFieldLocked(run) });
		else if (run.fieldCode !== undefined && top && top.separate === undefined) {
			top.code += run.fieldCode;
			top.codeRun ??= index;
			if (top.nameRun === undefined && /\S/.test(run.fieldCode)) top.nameRun = index;
		} else if (run.fieldChar === 'separate' && top) top.separate = index;
		else if (run.fieldChar === 'end' && top) {
			open.pop();
			const instr = top.code.trim();
			const names = {
				...(top.codeRun !== undefined ? { code: top.codeRun } : {}),
				...(top.nameRun !== undefined ? { name: top.nameRun } : {}),
			};
			if (top.separate === undefined) {
				if (options.resultless)
					found.push({
						from: top.begin,
						to: index,
						instr,
						locked: top.locked,
						runs: [],
						...names,
						insertAt: index,
						resultless: true,
					});
				return;
			}
			const results: number[] = [];
			let gone = false;
			for (let at = top.separate + 1; at < index; at++) {
				const item = runs[at]!;
				if (item.field?.instr !== instr || item.fieldChar || item.fieldCode !== undefined) continue;
				if (removed(item)) gone = true;
				else results.push(at);
			}
			found.push({
				from: top.begin,
				to: index,
				instr,
				locked: top.locked || results.some((at) => isFieldLocked(runs[at])),
				runs: results,
				...names,
				insertAt: top.separate + 1,
				...(gone && !results.length ? { removed: true as const } : {}),
			});
		}
	});
	let simple: FieldResultSpan | undefined;
	let identity: string | undefined;
	runs.forEach((run, index) => {
		if (!run.field?.simple) {
			simple = undefined;
			return;
		}
		if (simple && simple.instr === run.field.instr && identity === run.fieldInstanceId) {
			if (!removed(run)) simple.runs.push(index);
			simple.to = index;
		} else {
			simple = {
				from: index,
				to: index,
				instr: run.field.instr,
				locked: false,
				runs: removed(run) ? [] : [index],
				insertAt: index,
			};
			identity = run.fieldInstanceId;
			found.push(simple);
		}
		simple.locked ||= isFieldLocked(run);
	});
	return found;
}

/** Run properties that make a run something other than formatted text. */
const STRUCTURE = [
	'fieldChar',
	'fieldCode',
	'fieldInstanceId',
	'fieldFlags',
	'break',
	'noteReference',
	'noteMark',
	'noteSeparator',
	'equation',
	'image',
] as const;

/**
 * Run properties that are not character formatting: what a run holds, the revision it is part of,
 * its comments and its link. A result formatted like the field code keeps its own.
 */
const NOT_FORMATTING = [
	...STRUCTURE,
	'text',
	'field',
	'revision',
	'formatRevision',
	'restoredRunPropertiesXml',
	'commentIds',
	'link',
	'contentControl',
	'outerContentControls',
] as const;

/**
 * `result` with `text`, formatted like the field-code run `code` (Word's `\* CHARFORMAT` takes
 * the formatting of the first letter of the field name).
 */
export function withFieldCodeFormatting(result: TextRun, code: TextRun, text: string): TextRun {
	const formatted: Partial<TextRun> = { ...code };
	const own: Partial<TextRun> = {};
	for (const key of NOT_FORMATTING) {
		delete formatted[key];
		if (result[key] !== undefined) Object.assign(own, { [key]: result[key] });
	}
	return { ...formatted, ...own, text };
}

/** Whether two runs hold the same properties, in any order. */
function sameRun(a: TextRun, b: TextRun): boolean {
	const keys = new Set([...Object.keys(a), ...Object.keys(b)]) as Set<keyof TextRun>;
	return [...keys].every((key) => JSON.stringify(a[key]) === JSON.stringify(b[key]));
}

/** The result run `text` of a complex field that has none, formatted like its instruction run. */
export function fieldResultRun(code: TextRun, instr: string, text: string): TextRun {
	const formatted: Partial<TextRun> = { ...code };
	for (const key of STRUCTURE) delete formatted[key];
	return { ...formatted, text, field: { instr } };
}

/**
 * The runs with `field`'s result replaced by `text`, or undefined when it already reads `text`.
 * The text goes into the first result run, which keeps its formatting (Word's `\* MERGEFORMAT`),
 * or with `charFormat` takes that of the first letter of the field name, and the other result
 * runs are emptied rather than removed: the writer keeps every run where it was, with any
 * formatting the model does not hold. An empty result gets a run formatted like the instruction;
 * a field saved without a result slot (`resultless`) is left as it is.
 */
export function replaceFieldResult(
	runs: readonly TextRun[],
	field: FieldResultSpan,
	text: string,
	charFormat: boolean,
): TextRun[] | undefined {
	if (field.resultless) return undefined;
	const current = field.runs.map((at) => runs[at]!.text).join('') === text;
	const next = [...runs];
	const [first, ...rest] = field.runs;
	const name = field.name ?? field.code;
	if (first !== undefined) {
		const formatted =
			charFormat && name !== undefined
				? withFieldCodeFormatting(runs[first]!, runs[name]!, text)
				: undefined;
		// CHARFORMAT also brings the formatting up to date when the text already is.
		if (
			current &&
			(!formatted || (sameRun(formatted, runs[first]!) && rest.every((at) => !runs[at]!.text)))
		)
			return undefined;
		next[first] = formatted ?? { ...runs[first]!, text };
		for (const at of rest) next[at] = { ...runs[at]!, text: '' };
		return next;
	}
	const code = charFormat ? name : field.code;
	if (current || code === undefined || field.removed) return undefined;
	next.splice(field.insertAt, 0, fieldResultRun(runs[code]!, field.instr, text));
	return next;
}

/** A field's new result: its text, and whether it takes the field name's formatting. */
export interface FieldUpdate {
	text: string;
	charFormat?: boolean;
}

/** Gives an unlocked field its result, or undefined to keep the one it has. */
export type FieldEvaluator = (
	field: FieldResultSpan,
	/** The paragraph the field is in, as it was before any of its fields were updated. */
	paragraph: Paragraph,
) => FieldUpdate | undefined;

function paragraphWithResults(paragraph: Paragraph, evaluate: FieldEvaluator): Paragraph {
	let runs: readonly TextRun[] = paragraph.runs;
	// From the last field back, so an edit never moves the runs of a field still to be done.
	const start = (field: FieldResultSpan) => field.runs[0] ?? field.insertAt;
	for (const field of fieldResultSpans(runs).sort((a, b) => start(b) - start(a))) {
		const update = field.locked ? undefined : evaluate(field, paragraph);
		if (update)
			runs = replaceFieldResult(runs, field, update.text, update.charFormat === true) ?? runs;
	}
	return runs === paragraph.runs ? paragraph : { ...paragraph, runs: [...runs] };
}

/**
 * A table with its cells' results set, in the tables nested in a cell too (they are the cell's
 * source blocks; its own paragraphs are `paragraphs`). The table itself when nothing changed.
 */
function tableWithResults(table: Table, evaluate: FieldEvaluator): Table {
	let changed = false;
	const rows = table.rows.map((row) =>
		row.map((cell) => {
			const paragraphs = cell.paragraphs.map((item) => paragraphWithResults(item, evaluate));
			const sourceBlocks = cell.sourceBlocks?.map((item) =>
				item.type === 'table' ? tableWithResults(item, evaluate) : item,
			);
			if (
				paragraphs.every((item, index) => item === cell.paragraphs[index]) &&
				(sourceBlocks ?? []).every((item, index) => item === cell.sourceBlocks![index])
			)
				return cell;
			changed = true;
			return { ...cell, paragraphs, ...(sourceBlocks ? { sourceBlocks } : {}) };
		}),
	);
	return changed ? { ...table, rows } : table;
}

/**
 * The blocks with every unlocked field result that `evaluate` answers set to its answer.
 * Unchanged paragraphs and tables are returned as they are, so the writer keeps their XML.
 */
export function withFieldResults(blocks: readonly Block[], evaluate: FieldEvaluator): Block[] {
	let changed = false;
	const next = blocks.map((block): Block => {
		const updated =
			block.type === 'paragraph'
				? paragraphWithResults(block, evaluate)
				: tableWithResults(block, evaluate);
		changed ||= updated !== block;
		return updated;
	});
	return changed ? next : (blocks as Block[]);
}
