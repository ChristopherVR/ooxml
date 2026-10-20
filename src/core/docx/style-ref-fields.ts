// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Brings the cached results of STYLEREF fields up to date in a story's paragraphs.
import { isFieldLocked } from './field-lock';
import type { Block, Paragraph, TextRun } from './model';
import { parseStyleReference, type StyleReference } from './style-ref';

/** A field's cached result in a paragraph's runs. */
interface FieldResult {
	instr: string;
	locked: boolean;
	/** The runs holding the result. */
	runs: number[];
	/** The run holding the instruction of a complex field. */
	code?: number;
	/** Where a result goes when there is none: just after the `separate` marker. */
	insertAt: number;
}

/** Complex fields closed in this paragraph, and simple fields, with the runs of their results. */
function fieldResults(runs: readonly TextRun[]): FieldResult[] {
	const found: FieldResult[] = [];
	const open: { code: string; codeRun?: number; separate?: number; locked: boolean }[] = [];
	runs.forEach((run, index) => {
		const top = open.at(-1);
		if (run.fieldChar === 'begin') open.push({ code: '', locked: isFieldLocked(run) });
		else if (run.fieldCode !== undefined && top && top.separate === undefined) {
			top.code += run.fieldCode;
			top.codeRun ??= index;
		} else if (run.fieldChar === 'separate' && top) top.separate = index;
		else if (run.fieldChar === 'end' && top) {
			open.pop();
			if (top.separate === undefined) return;
			const instr = top.code.trim();
			const results: number[] = [];
			for (let at = top.separate + 1; at < index; at++) {
				const item = runs[at]!;
				if (item.field?.instr === instr && !item.fieldChar && item.fieldCode === undefined)
					results.push(at);
			}
			found.push({
				instr,
				locked: top.locked || results.some((at) => isFieldLocked(runs[at])),
				runs: results,
				...(top.codeRun !== undefined ? { code: top.codeRun } : {}),
				insertAt: top.separate + 1,
			});
		}
	});
	let simple: FieldResult | undefined;
	let identity: string | undefined;
	runs.forEach((run, index) => {
		if (!run.field?.simple) {
			simple = undefined;
			return;
		}
		if (simple && simple.instr === run.field.instr && identity === run.fieldInstanceId)
			simple.runs.push(index);
		else {
			simple = { instr: run.field.instr, locked: false, runs: [index], insertAt: index };
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
 * The runs with `field`'s result replaced by `text`, or undefined when it already reads `text`.
 * The text goes into the first result run, which keeps its formatting (Word's `\* MERGEFORMAT`),
 * and the other result runs are emptied rather than removed: the writer keeps every run where it
 * was, with any formatting the model does not hold. An empty result gets a run formatted like the
 * instruction.
 */
function replaceResult(runs: readonly TextRun[], field: FieldResult, text: string) {
	if (field.runs.map((at) => runs[at]!.text).join('') === text) return undefined;
	const next = [...runs];
	const [first, ...rest] = field.runs;
	if (first !== undefined) {
		next[first] = { ...runs[first]!, text };
		for (const at of rest) next[at] = { ...runs[at]!, text: '' };
		return next;
	}
	if (field.code === undefined) return undefined;
	const formatted: Partial<TextRun> = { ...runs[field.code]! };
	for (const key of STRUCTURE) delete formatted[key];
	next.splice(field.insertAt, 0, { ...formatted, text, field: { instr: field.instr } });
	return next;
}

/** Gives a STYLEREF field its result, or undefined to keep the one it has. */
export type StyleRefEvaluator = (
	reference: StyleReference,
	paragraph: Paragraph,
) => string | undefined;

function paragraphWithResults(paragraph: Paragraph, evaluate: StyleRefEvaluator): Paragraph {
	let runs: readonly TextRun[] = paragraph.runs;
	// From the last field back, so an edit never moves the runs of a field still to be done.
	const start = (field: FieldResult) => field.runs[0] ?? field.insertAt;
	for (const field of fieldResults(runs).sort((a, b) => start(b) - start(a))) {
		const reference = field.locked ? undefined : parseStyleReference(field.instr);
		const text = reference && evaluate(reference, paragraph);
		if (text !== undefined) runs = replaceResult(runs, field, text) ?? runs;
	}
	return runs === paragraph.runs ? paragraph : { ...paragraph, runs: [...runs] };
}

/**
 * The blocks with every unlocked STYLEREF result that `evaluate` answers set to its answer.
 * Unchanged paragraphs and tables are returned as they are, so the writer keeps their XML.
 */
export function withStyleRefResults(
	blocks: readonly Block[],
	evaluate: StyleRefEvaluator,
): Block[] {
	let changed = false;
	const next = blocks.map((block): Block => {
		if (block.type === 'paragraph') {
			const paragraph = paragraphWithResults(block, evaluate);
			changed ||= paragraph !== block;
			return paragraph;
		}
		let tableChanged = false;
		const rows = block.rows.map((row) =>
			row.map((cell) => {
				const paragraphs = cell.paragraphs.map((item) => paragraphWithResults(item, evaluate));
				if (paragraphs.every((item, index) => item === cell.paragraphs[index])) return cell;
				tableChanged = true;
				return { ...cell, paragraphs };
			}),
		);
		if (!tableChanged) return block;
		changed = true;
		return { ...block, rows };
	});
	return changed ? next : (blocks as Block[]);
}
