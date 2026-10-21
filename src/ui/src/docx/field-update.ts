import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	bodyStyleRefResult,
	fieldResultRun,
	isFieldLocked,
	parseStyleReference,
	styleRefSource,
	withFieldCodeFormatting,
	type DocumentModel,
	type Paragraph,
	type StyleRefSource,
} from 'ooxml-core/docx';
import { reviewSourceLength } from 'ooxml-core/docx/layout';
import {
	emptySimpleFieldResult,
	fieldResultRanges,
	inlineNodeRun,
	runToInlineNodes,
	type FieldResultRange,
} from 'ooxml-core/docx/ui';
import type { Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { renumberCaptions, seqLabelOf } from './caption-commands';
import { docToModel } from './model-adapter';
import { schema } from './schema';

const REFERENCE = /^\s*(REF|PAGEREF)\s+("[^"]+"|\S+)(.*)$/i;

/**
 * Switches that do not change a REF/PAGEREF result (`\h` hyperlink, `\* MERGEFORMAT`). Any other
 * switch (`\n`, `\r`, `\w`, `\p`, `\f`...) changes what the field shows, so those fields are left
 * alone rather than being given a wrong result.
 */
function plainSwitches(rest: string): boolean {
	return !/\\/.test(rest.replace(/\\h\b/gi, '').replace(/\\\*\s+\S+/g, ''));
}

/** The paragraph a bookmark sits on, with its id and text. */
function bookmarkTargets(doc: ProseMirrorNode): Map<string, ProseMirrorNode> {
	const targets = new Map<string, ProseMirrorNode>();
	doc.descendants((node) => {
		if (node.type.name !== 'paragraph') return true;
		for (const name of (node.attrs.bookmarks as string[]) ?? [])
			if (name && !targets.has(name)) targets.set(name, node);
		return false;
	});
	return targets;
}

/** "Figure 3": the caption text up to and including its SEQ number, or undefined without one. */
function captionLabel(paragraph: ProseMirrorNode): string | undefined {
	const text = paragraph.textContent;
	let number: string | undefined;
	paragraph.forEach((child) => {
		if (seqLabelOf(child) !== undefined) number = child.text ?? '';
	});
	return number === undefined
		? undefined
		: text.slice(0, text.indexOf(number) + number.length).trim();
}

/**
 * The instruction marker holding the first letter of the field name, for the result of a
 * complex field that starts at child `index` of `paragraph`; undefined for a simple field.
 */
function fieldNameMarker(paragraph: ProseMirrorNode, index: number): ProseMirrorNode | undefined {
	const marker = (at: number) => {
		const child = paragraph.maybeChild(at);
		return child?.type.name === 'fieldMarker' ? child : undefined;
	};
	if (marker(index - 1)?.attrs.kind !== 'separate') return undefined;
	let name: ProseMirrorNode | undefined;
	for (let at = index - 2; at >= 0; at--) {
		const child = marker(at);
		if (child?.attrs.kind === 'begin') return name;
		if (child?.attrs.kind !== 'code') return undefined;
		if (/\S/.test(String(child.attrs.code ?? ''))) name = child;
	}
	return undefined;
}

/** A complex field whose result is empty, in one paragraph. */
interface EmptyField {
	/** Where its result goes: just after its `separate` marker. */
	at: number;
	instr: string;
	paragraph: ProseMirrorNode;
	begin: ProseMirrorNode;
	/** Its first instruction marker, and the one holding the first letter of the field name. */
	code?: ProseMirrorNode;
	name?: ProseMirrorNode;
	/** Where its markers start and end in the paragraph. */
	from: number;
	to: number;
}

/** The complex fields of `doc` that begin and end in one paragraph with nothing as their result. */
function emptyFields(doc: ProseMirrorNode): EmptyField[] {
	const found: EmptyField[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		const open: (Omit<EmptyField, 'at' | 'to'> & { separate?: number })[] = [];
		node.forEach((child, offset) => {
			if (child.type.name !== 'fieldMarker') return;
			const top = open.at(-1);
			const kind = child.attrs.kind;
			if (kind === 'begin') open.push({ instr: '', paragraph: node, begin: child, from: offset });
			else if (kind === 'code' && top && top.separate === undefined) {
				const code = String(child.attrs.code ?? '');
				top.instr += code;
				top.code ??= child;
				if (!top.name && /\S/.test(code)) top.name = child;
			} else if (kind === 'separate' && top) top.separate = offset + child.nodeSize;
			else if (kind === 'end' && top) {
				open.pop();
				const { separate, ...field } = top;
				if (separate === offset)
					found.push({ ...field, instr: field.instr.trim(), at: pos + 1 + offset, to: offset + 1 });
			}
		});
		return false;
	});
	return found;
}

/**
 * The result runs for the empty STYLEREF fields of `doc`: each takes the nearest text in its
 * style, formatted like its instruction (with `\* CHARFORMAT`, like the first letter of its name).
 */
function emptyStyleRefResults(
	doc: ProseMirrorNode,
	source: StyleRefSource,
): { at: number; nodes: ProseMirrorNode[] }[] {
	const results: { at: number; nodes: ProseMirrorNode[] }[] = [];
	for (const field of emptyFields(doc)) {
		const reference = parseStyleReference(field.instr);
		if (!reference || isFieldLocked(inlineNodeRun(field.begin))) continue;
		const at = source.index.get(String(field.paragraph.attrs.id));
		const paragraph = at === undefined ? undefined : source.paragraphs[at];
		const own = paragraph && runsAt(paragraph, field.from, field.to);
		const text = paragraph && bodyStyleRefResult(source, reference, paragraph, own);
		const code = reference.charFormat ? (field.name ?? field.code) : field.code;
		const run = code && inlineNodeRun(code);
		if (text && run)
			results.push({
				at: field.at,
				nodes: runToInlineNodes(fieldResultRun(run, field.instr, text), schema),
			});
	}
	return results;
}

/** The runs of a model paragraph that hold its editor positions `start` to `end`. */
function runsAt(
	paragraph: Paragraph,
	start: number,
	end: number,
): { from: number; to: number } | undefined {
	let offset = 0;
	let from: number | undefined;
	let to: number | undefined;
	paragraph.runs.forEach((run, index) => {
		const next = offset + reviewSourceLength(run);
		if (from === undefined && next > start) from = index;
		if (offset < end) to = index;
		offset = next;
	});
	return from === undefined || to === undefined ? undefined : { from, to };
}

/**
 * Word's Update Field for the whole body: every `SEQ` caption is renumbered, then `REF` results
 * follow their bookmarked paragraph and `PAGEREF` results take the page from `pageOf`. A REF that
 * pointed at a caption's label and number keeps showing the label; a longer result is treated as
 * the entire caption. With the document's `model`, `STYLEREF` results take the nearest text in
 * their style before the field, else after it, which an empty paragraph in the style makes empty
 * (the field stays, with its instruction, and a later update fills it). Fields with other
 * switches or a missing bookmark or style are left as they are. One undoable step; returns
 * whether anything changed.
 */
export function updateFields(
	view: EditorView,
	pageOf: (id: string) => string | undefined,
	model?: DocumentModel,
): boolean {
	if (!view.editable) return false;
	const tr = view.state.tr;
	refreshFieldResults(tr, pageOf, false, model);
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}

/**
 * The refresh behind `updateFields`, applied to `tr` so other commands (a new caption) can include
 * it in their own undo step. Without `pageOf`, `PAGEREF` results are left alone: pages need a layout.
 * With `captionsOnly`, only references that quote a caption are refreshed.
 */
export function refreshFieldResults(
	tr: Transaction,
	pageOf?: (id: string) => string | undefined,
	captionsOnly = false,
	model?: DocumentModel,
): Transaction {
	for (const label of new Set(
		fieldResultRanges(tr.doc)
			.map((run) => seqLabelOf(schema.text('x', [run.mark])))
			.filter((label): label is string => label !== undefined)
			.map((label) => label.toLowerCase()),
	))
		renumberCaptions(tr, label);
	const targets = bookmarkTargets(tr.doc);
	// STYLEREF reads the document as the captions renumbered above have left it.
	const current =
		model && !captionsOnly && tr.docChanged ? docToModel(tr.doc, model, tr.before) : model;
	const source = current && !captionsOnly ? styleRefSource(current) : undefined;
	const changes: Array<{ run: FieldResultRange; text: string; code?: ProseMirrorNode }> = [];
	for (const run of fieldResultRanges(tr.doc)) {
		if (isFieldLocked(inlineNodeRun(schema.text('x', run.marks)))) continue;
		const styleRef = source && parseStyleReference(String(run.mark.attrs.instr ?? ''));
		if (styleRef) {
			const $from = tr.doc.resolve(run.from);
			const at = source.index.get(String($from.parent.attrs.id));
			const paragraph = at === undefined ? undefined : source.paragraphs[at];
			const field =
				paragraph && runsAt(paragraph, run.from - $from.start(), run.to - $from.start());
			const text = paragraph && bodyStyleRefResult(source, styleRef, paragraph, field);
			// `\* CHARFORMAT`: the result takes the formatting of the field name's first letter.
			const code = styleRef.charFormat ? fieldNameMarker($from.parent, $from.index()) : undefined;
			if (text !== undefined && text !== run.text)
				changes.push({ run, text, ...(code ? { code } : {}) });
			continue;
		}
		const match = REFERENCE.exec(String(run.mark.attrs.instr ?? ''));
		if (!match || !plainSwitches(match[3] ?? '')) continue;
		const target = targets.get((match[2] ?? '').replace(/^"|"$/g, ''));
		if (!target || (captionsOnly && captionLabel(target) === undefined)) continue;
		let text: string | undefined;
		if (match[1]!.toUpperCase() === 'PAGEREF') text = pageOf?.(String(target.attrs.id));
		else {
			const label = captionLabel(target);
			const whole = target.textContent.trim();
			text = label !== undefined && run.text.trim().length <= label.length ? label : whole;
		}
		if (text && text !== run.text) changes.push({ run, text });
	}
	const edits = changes.map(({ run, text, code }) => ({
		at: run.from,
		apply: () => {
			// An empty result keeps its field: a simple one becomes a complex field with no result.
			if (!text) {
				if (!emptySimpleFieldResult(tr, run.from, run.to)) tr.delete(run.from, run.to);
				return;
			}
			const result = code && inlineNodeRun(tr.doc.nodeAt(run.from)!);
			const codeRun = code && inlineNodeRun(code);
			tr.replaceWith(
				run.from,
				run.to,
				result && codeRun
					? runToInlineNodes(withFieldCodeFormatting(result, codeRun, text), schema)
					: schema.text(text, run.marks),
			);
		},
	}));
	if (source)
		for (const { at, nodes } of emptyStyleRefResults(tr.doc, source))
			edits.push({ at, apply: () => void tr.insert(at, nodes) });
	// Later fields first, so an edit never shifts the positions still to be processed.
	for (const edit of edits.sort((a, b) => b.at - a.at)) edit.apply();
	return tr;
}
