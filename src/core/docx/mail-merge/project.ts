import { mapStoryBlocks } from '../cell-blocks';
import { computedFields, type ComputedFieldResult } from '../computed-fields';
import {
	fieldResultSpans,
	withFieldCodeFormatting,
	type FieldResultSpan,
} from '../field-result-spans';
import { fieldName } from '../field-runs';
import { withStoryBlocks } from '../field-stories';
import type { Block, DocumentModel, Paragraph, TextRun } from '../model';
import type { MailMergeData } from './data-source';
import { isMergeField, MergeFields, type MailMergeOptions } from './field-values';

interface FieldNode {
	span: FieldResultSpan;
	children: FieldNode[];
	containsMerge: boolean;
}

function fieldTree(runs: readonly TextRun[]): FieldNode[] {
	const roots: FieldNode[] = [];
	const open: FieldNode[] = [];
	for (const span of fieldResultSpans(runs, { resultless: true }).sort(
		(a, b) => a.from - b.from || b.to - a.to,
	)) {
		while (open.length && open.at(-1)!.span.to < span.from) {
			const closed = open.pop()!;
			if (closed.containsMerge && open.length) open.at(-1)!.containsMerge = true;
		}
		const parent = open.at(-1);
		const node = { span, children: [], containsMerge: isMergeField(fieldName(span.instr)) };
		(parent?.children ?? roots).push(node);
		open.push(node);
	}
	while (open.length > 1) {
		const closed = open.pop()!;
		if (closed.containsMerge) open.at(-1)!.containsMerge = true;
	}
	return roots;
}

function cachedText(runs: readonly TextRun[], span: FieldResultSpan): string {
	return span.runs.map((index) => runs[index]?.text ?? '').join('');
}

function resultRun(
	paragraph: Paragraph,
	span: FieldResultSpan,
	text: string,
	charFormat: boolean,
): TextRun {
	const code = paragraph.runs[span.name ?? span.code ?? span.from];
	const source = paragraph.runs[span.runs[0] ?? -1] ?? code ?? { text: '' };
	const formatted = charFormat && code ? withFieldCodeFormatting(source, code, text) : source;
	const {
		field: _field,
		fieldInstanceId: _fieldInstanceId,
		fieldChar: _fieldChar,
		fieldCode: _fieldCode,
		fieldFlags: _fieldFlags,
		...run
	} = formatted;
	return { ...run, text: text.replace(/\r\n?/g, '\n') };
}

function projectParagraph(
	paragraph: Paragraph,
	fields: MergeFields,
	computed: ReturnType<typeof computedFields>,
	body: boolean,
	story: { openFields: number; index: number },
): Paragraph {
	if (fields.skipped) return paragraph;
	const runs = paragraph.runs;
	// A cross-paragraph field can enclose locally balanced fields. Preserve that whole range.
	let first = story.openFields ? runs.length : 0;
	let last = first;
	for (const [index, run] of runs.entries()) {
		if (run.fieldChar === 'begin') story.openFields++;
		if (run.fieldChar === 'end' && story.openFields > 0) {
			story.openFields--;
			if (story.openFields === 0 && first === runs.length) first = index + 1;
		}
		if (!story.openFields) last = index + 1;
	}
	const roots = fieldTree(runs).filter(({ span }) => span.from >= first && span.to < last);
	const evaluate = (
		node: FieldNode,
		depth: number,
		nested: boolean,
	): ComputedFieldResult | undefined => {
		const { span } = node;
		if (span.locked || span.removed) return undefined;
		if (depth > 64) return fields.issue('invalid-instruction', paragraph.id, fieldName(span.instr));
		const codeEnd = span.insertAt - (span.resultless ? 0 : 1);
		const codeChildren = node.children.filter((child) => child.span.to < codeEnd);
		const codeNodes = new Map(
			codeChildren.map((child, index) => [child.span.from, { child, index }]),
		);
		let instruction = span.instr;
		if (codeNodes.size) {
			instruction = '';
			for (let at = span.from + 1; at < span.insertAt; at++) {
				const entry = codeNodes.get(at);
				if (entry) {
					instruction += `\uE000${entry.index}\uE001`;
					at = entry.child.span.to;
				} else instruction += runs[at]?.fieldCode ?? '';
			}
		}
		const resolved = new Map<number, string>();
		const result = fields.evaluate({
			instruction,
			blockId: paragraph.id,
			cached: cachedText(runs, span),
			allowRecordRules: body && !nested,
			nestedMerge: codeChildren.some((child) => child.containsMerge),
			position: `${story.index}:${span.from}`,
			resolve: (arg) =>
				arg.replace(/\uE000(\d+)\uE001/g, (placeholder: string, index: string) => {
					const id = Number(index);
					const child = codeChildren[id];
					if (!child) return placeholder;
					const old = resolved.get(id);
					if (old !== undefined) return old;
					const value = evaluate(child, depth + 1, true)?.text ?? cachedText(runs, child.span);
					resolved.set(id, value);
					return value;
				}),
		});
		return result ?? (nested && !node.containsMerge ? computed.result(span, paragraph) : undefined);
	};
	const next: TextRun[] = [];
	let cursor = 0;
	let changed = false;
	const visit = (nodes: FieldNode[]) => {
		for (const node of nodes) {
			if (fields.skipped) break;
			const { span } = node;
			const result = evaluate(node, 0, false);
			if (result) {
				next.push(...runs.slice(cursor, span.from));
				if (result.text)
					next.push(resultRun(paragraph, span, result.text, result.charFormat === true));
				cursor = span.to + 1;
				changed = true;
			} else if (!span.locked && !span.removed) {
				// A saved result can itself contain a merge field. Its outer instruction stays intact.
				visit(node.children.filter((child) => child.span.from >= span.insertAt));
			}
		}
	};
	visit(roots);
	let rootIndex = 0;
	for (const [index, run] of runs.entries()) {
		while (roots[rootIndex] && roots[rootIndex]!.span.to < index) rootIndex++;
		const root = roots[rootIndex];
		if (run.fieldCode && (!root || index < root.span.from)) {
			const name = fieldName(run.fieldCode);
			if (isMergeField(name)) fields.issue('invalid-instruction', paragraph.id, name);
		}
	}
	if (!changed) return paragraph;
	next.push(...runs.slice(cursor));
	return { ...paragraph, runs: next.length ? next : [{ text: '' }] };
}

export function projectMerge(
	model: DocumentModel,
	data: MailMergeData,
	start: number,
	sequence: number,
	options: MailMergeOptions,
) {
	const fields = new MergeFields(data, start, sequence, options);
	const computed = computedFields(model);
	let storyIndex = 0;
	const project = (blocks: Block[], body: boolean) => {
		const story = { openFields: 0, index: storyIndex++ };
		const paragraph = (value: Paragraph) => projectParagraph(value, fields, computed, body, story);
		const table = blocks[0];
		if (!body || model.mailMerge?.mainDocumentType !== 'mailingLabels' || table?.type !== 'table')
			return mapStoryBlocks(blocks, paragraph);
		const rows = table.rows.map((row) =>
			row.map((cell) => {
				if (cell.sourceBlocks) return cell;
				const paragraphs = cell.paragraphs.map(paragraph);
				if (!fields.exhausted) return { ...cell, paragraphs };
				return { ...cell, paragraphs: paragraphs.map((p) => ({ ...p, runs: [{ text: '' }] })) };
			}),
		);
		return [{ ...table, rows }, ...mapStoryBlocks(blocks.slice(1), paragraph)];
	};
	const blocks = project(model.blocks, true);
	const nextIndex = fields.index + 1;
	fields.index = start;
	const stories = { ...model, blocks };
	const projected = withStoryBlocks(stories, (blocks) => project([...blocks], false));
	for (const key of ['footnotes', 'endnotes'] as const) {
		const notes = projected[key];
		if (notes)
			projected[key] = notes.map((note) => ({ ...note, blocks: project(note.blocks, false) }));
	}
	return { model: projected, skipped: fields.skipped, nextIndex, issues: fields.issues };
}
