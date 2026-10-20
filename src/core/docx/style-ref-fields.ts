// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Brings the cached results of STYLEREF fields up to date in a story's paragraphs.
import { isFieldLocked } from './field-lock';
import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	HeaderFooterSlots,
	Paragraph,
	TextRun,
} from './model';
import {
	parseStyleReference,
	styleRefSource,
	type RunSpan,
	type StyleRefSource,
	type StyleReference,
} from './style-ref';

/** A field's cached result in a paragraph's runs; `from` and `to` span the whole field. */
interface FieldResult extends RunSpan {
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
	const open: {
		begin: number;
		code: string;
		codeRun?: number;
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
				from: top.begin,
				to: index,
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
		if (simple && simple.instr === run.field.instr && identity === run.fieldInstanceId) {
			simple.runs.push(index);
			simple.to = index;
		} else {
			simple = {
				from: index,
				to: index,
				instr: run.field.instr,
				locked: false,
				runs: [index],
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
	/** The field's own runs in `paragraph`. */
	field: RunSpan,
) => string | undefined;

function paragraphWithResults(paragraph: Paragraph, evaluate: StyleRefEvaluator): Paragraph {
	let runs: readonly TextRun[] = paragraph.runs;
	// From the last field back, so an edit never moves the runs of a field still to be done.
	const start = (field: FieldResult) => field.runs[0] ?? field.insertAt;
	for (const field of fieldResults(runs).sort((a, b) => start(b) - start(a))) {
		const reference = field.locked ? undefined : parseStyleReference(field.instr);
		const text = reference && evaluate(reference, paragraph, field);
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

/**
 * The result a header or footer STYLEREF shows on every page, when that does not depend on the
 * page: every text in the document it could pick up reads the same (typically the style is used
 * once, for a title or a document number). Undefined otherwise, or when nothing is in the style.
 */
function pageIndependentResult(
	source: StyleRefSource,
	reference: StyleReference,
): string | undefined {
	const target = source.target(reference.style);
	if (!target) return undefined;
	let result: string | undefined;
	for (let index = 0; index < source.paragraphs.length; index++)
		for (const text of source.results(index, reference, target)) {
			if (result !== undefined && text !== result) return undefined;
			result = text;
		}
	return result;
}

/**
 * The result of a STYLEREF in the body: the nearest text in the style before the field, else the
 * nearest after it (ECMA-376 Part 1, 17.16.5; Word searches back from the field, then on from
 * it). Text in a character style in the field's own paragraph is nearest, the field's own runs
 * (`field`) left out; a paragraph style's match is a whole paragraph, so the field's own is not
 * one. Undefined when the paragraph is not in the main document or nothing is in the style.
 */
export function bodyStyleRefResult(
	source: StyleRefSource,
	reference: StyleReference,
	paragraph: Paragraph,
	field?: RunSpan,
): string | undefined {
	const target = source.target(reference.style);
	const at = source.index.get(paragraph.id);
	if (!target || at === undefined) return undefined;
	const own =
		field && target.kind === 'character' ? source.portions(at, reference, target, field) : [];
	const earlier = own.filter((portion) => portion.to < field!.from);
	if (earlier.length) return earlier.at(-1)!.text;
	for (let index = at - 1; index >= 0; index--) {
		const found = source.results(index, reference, target);
		if (found.length) return found.at(-1);
	}
	const later = own.find((portion) => portion.from > field!.to);
	if (later) return later.text;
	for (let index = at + 1; index < source.paragraphs.length; index++) {
		const found = source.results(index, reference, target);
		if (found.length) return found[0];
	}
	return undefined;
}

const SLOTS = ['default', 'first', 'even'] as const;

/**
 * Brings STYLEREF results up to date. A field in a header or footer takes the text it shows on
 * every page, when that does not depend on the page (`pageIndependentResult`); Word updates such
 * fields as the page is drawn, so a saved header should not keep a stale value. With
 * `body: true`, fields in the body take the nearest text in their style before them, else after
 * them, as Word's Update Field does. Locked fields and fields with switches that are not evaluated
 * keep their result. Returns `model` itself when nothing changed.
 */
export function refreshStyleRefFields(
	model: DocumentModel,
	options: { body?: boolean } = {},
): DocumentModel {
	let source: StyleRefSource | undefined;
	const read = () => (source ??= styleRefSource(model));
	const everywhere = new Map<string, string | undefined>();
	const story: StyleRefEvaluator = (reference) => {
		const key = JSON.stringify(reference);
		if (!everywhere.has(key)) everywhere.set(key, pageIndependentResult(read(), reference));
		return everywhere.get(key);
	};
	const content = (item: HeaderFooterContent | undefined) => {
		if (!item) return item;
		const blocks = withStyleRefResults(item.blocks, story);
		return blocks === item.blocks ? item : { ...item, blocks };
	};
	const slots = (item: HeaderFooterSlots | undefined) => {
		if (!item) return item;
		let changed = false;
		const next: HeaderFooterSlots = { ...item };
		for (const slot of SLOTS) {
			const refreshed = content(item[slot]);
			if (refreshed === item[slot]) continue;
			changed = true;
			next[slot] = refreshed!;
		}
		return changed ? next : item;
	};
	let changed = false;
	const sections = model.sections?.map((section) => {
		const headers = slots(section.headers);
		const footers = slots(section.footers);
		if (headers === section.headers && footers === section.footers) return section;
		changed = true;
		return {
			...section,
			...(headers ? { headers } : {}),
			...(footers ? { footers } : {}),
		};
	});
	const blocks = options.body
		? withStyleRefResults(model.blocks, (reference, paragraph, field) =>
				bodyStyleRefResult(read(), reference, paragraph, field),
			)
		: model.blocks;
	if (!changed && blocks === model.blocks) return model;
	return { ...model, blocks, ...(sections ? { sections } : {}) };
}
