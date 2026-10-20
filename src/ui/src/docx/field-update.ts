import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	bodyStyleRefResult,
	isFieldLocked,
	parseStyleReference,
	styleRefSource,
	type DocumentModel,
	type Paragraph,
} from 'ooxml-core/docx';
import { reviewSourceLength } from 'ooxml-core/docx/layout';
import { fieldResultRanges, inlineNodeRun, type FieldResultRange } from 'ooxml-core/docx/ui';
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
 * their style before the field, else after it. Fields with other switches or a missing bookmark
 * or style are left as they are. One undoable step; returns whether anything changed.
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
	const changes: Array<{ run: FieldResultRange; text: string }> = [];
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
			if (text && text !== run.text) changes.push({ run, text });
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
	// Later fields first, so replacing one never shifts the positions still to be processed.
	for (const { run, text } of changes.reverse())
		tr.replaceWith(run.from, run.to, schema.text(text, run.marks));
	return tr;
}
