import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { fieldClipboardSlice } from './field-clipboard';
import { inlineNodeRun } from './run-adapter';
import type { ExtraRunProperties } from './run-extra-mark';
import {
	deleteSimpleFieldResult,
	replaceSimpleFieldResult,
	simpleFieldPasteSlice,
} from './simple-field-input';

/** Derived field metadata repairs are not user formatting edits. */
export const FIELD_RESULT_REPAIR_META = 'dve-field-result-repair';

/** Field marker kinds in document order. */
function markerKinds(doc: ProseMirrorNode): string[] {
	const kinds: string[] = [];
	doc.descendants((node) => {
		if (node.type.name === 'fieldMarker') kinds.push(String(node.attrs.kind));
	});
	return kinds;
}

/**
 * Whether complex-field markers nest like Word expects: each `begin` is followed by its code, an
 * optional `separate` and a matching `end`. Fields may nest and may span paragraphs.
 */
export function fieldsBalanced(doc: ProseMirrorNode): boolean {
	const stack: ('code' | 'result')[] = [];
	for (const kind of markerKinds(doc)) {
		if (kind === 'begin') stack.push('code');
		else if (kind === 'code') {
			if (stack.at(-1) !== 'code') return false;
		} else if (kind === 'separate') {
			if (stack.at(-1) !== 'code') return false;
			stack[stack.length - 1] = 'result';
		} else if (kind === 'end') {
			if (!stack.length) return false;
			stack.pop();
		}
	}
	return stack.length === 0;
}

/**
 * Keeps complex fields structurally intact: an edit that would delete only part of a field (its
 * begin, separator or end marker) is rejected, while editing the result text or deleting a whole
 * field is allowed. Documents that arrive unbalanced are left alone.
 */
export function fieldGuardPlugin(): Plugin {
	return new Plugin({
		props: {
			handleDOMEvents: {
				cut: (view, event) => {
					if (!view.editable || !event.clipboardData) return false;
					const tr = deleteFieldSelection(view.state);
					if (!tr) return false;
					const { dom, text } = view.serializeForClipboard(view.state.selection.content());
					event.clipboardData.clearData();
					event.clipboardData.setData('text/html', dom.innerHTML);
					event.clipboardData.setData('text/plain', text);
					event.preventDefault();
					view.dispatch(tr.scrollIntoView().setMeta('uiEvent', 'cut'));
					return true;
				},
			},
			transformCopied: fieldClipboardSlice,
			transformPasted: (slice, view) =>
				simpleFieldPasteSlice(fieldClipboardSlice(slice), view?.state),
			handleTextInput: (view, from, to, text) => {
				const transaction =
					replaceAroundMarkers(view.state, from, to, text) ??
					replaceSimpleFieldResult(view.state, from, to, text);
				if (transaction) view.dispatch(transaction);
				return Boolean(transaction);
			},
			handleKeyDown: (view, event) => {
				if (event.key !== 'Backspace' && event.key !== 'Delete') return false;
				const transaction =
					deleteAroundMarkers(view.state, event.key === 'Backspace') ??
					deleteSimpleFieldResult(view.state, event.key === 'Backspace');
				if (transaction) view.dispatch(transaction);
				return Boolean(transaction);
			},
		},
		appendTransaction: (transactions, _old, state) =>
			transactions.some((transaction) => transaction.docChanged)
				? markFieldResults(state)
				: transactions.some((transaction) => transaction.selectionSet)
					? trimSelection(state)
					: null,
		filterTransaction(transaction, state) {
			if (!transaction.docChanged) return true;
			let hadMarkers = false;
			state.doc.descendants((node) => {
				if (node.type.name === 'fieldMarker') hadMarkers = true;
				return !hadMarkers;
			});
			if (!hadMarkers || !fieldsBalanced(state.doc)) return true;
			return fieldsBalanced(transaction.doc);
		},
	});
}

/** Field-aware selection deletion for clipboard hosts; null delegates ordinary deletion. */
export function deleteFieldSelection(state: EditorState): Transaction | null {
	const { from, to, empty } = state.selection;
	return empty
		? null
		: (replaceAroundMarkers(state, from, to, '') ?? replaceSimpleFieldResult(state, from, to, ''));
}

/**
 * Tags text typed into a complex field's result (between its separator and end) with the field
 * mark, so replacing a whole result keeps it a field result as it does in Word. The innermost
 * begin marker owns the result's lock state; cached run formatting remains independent.
 */
function markFieldResults(state: EditorState): Transaction | null {
	const fieldMark = state.schema.marks.field;
	if (!fieldMark) return null;
	const stack: { code: string; inResult: boolean; locked: boolean | undefined }[] = [];
	const transaction = state.tr;
	state.doc.descendants((node, pos) => {
		if (node.type.name === 'fieldMarker') {
			const kind = node.attrs.kind;
			if (kind === 'begin')
				stack.push({ code: '', inResult: false, locked: inlineNodeRun(node)?.fieldFlags?.locked });
			else if (kind === 'code' && stack.length) {
				const open = stack[stack.length - 1];
				if (open) open.code += node.attrs.code ?? '';
			} else if (kind === 'separate' && stack.length) {
				const open = stack[stack.length - 1];
				if (open) open.inResult = true;
			} else if (kind === 'end') stack.pop();
			return false;
		}
		const top = stack.at(-1);
		if (node.isText && top?.inResult) {
			const end = pos + node.nodeSize;
			if (!fieldMark.isInSet(node.marks))
				transaction.addMark(pos, end, fieldMark.create({ instr: top.code.trim(), simple: false }));
			const type = state.schema.marks.runProperties;
			const current = type?.isInSet(node.marks);
			const props = (current?.attrs.props ?? {}) as ExtraRunProperties;
			if (type && props.fieldFlags?.locked !== top.locked) {
				const next = { ...props };
				const flags = { ...props.fieldFlags };
				if (top.locked === undefined) delete flags.locked;
				else flags.locked = top.locked;
				if (Object.keys(flags).length) next.fieldFlags = flags;
				else delete next.fieldFlags;
				if (current) transaction.removeMark(pos, end, current);
				if (Object.keys(next).length)
					transaction.addMark(pos, end, type.create({ ...current?.attrs, props: next }));
			}
		}
		return true;
	});
	return transaction.docChanged ? transaction.setMeta(FIELD_RESULT_REPAIR_META, true) : null;
}

const isMarker = (node: ProseMirrorNode | null | undefined) => node?.type.name === 'fieldMarker';

/** Drops hidden field markers from a text selection's edges, so a selected result is just its text. */
function trimSelection(state: EditorState): Transaction | null {
	const { selection, doc } = state;
	if (!(selection instanceof TextSelection) || selection.empty) return null;
	let { from, to } = selection;
	while (from < to && isMarker(doc.resolve(from).nodeAfter)) from++;
	while (to > from && isMarker(doc.resolve(to).nodeBefore)) to--;
	if (from === selection.from && to === selection.to) return null;
	const reversed = selection.anchor > selection.head;
	return state.tr.setSelection(
		TextSelection.create(doc, reversed ? to : from, reversed ? from : to),
	);
}

function markersIn(
	doc: ProseMirrorNode,
	from: number,
	to: number,
): { pos: number; size: number }[] {
	const markers: { pos: number; size: number }[] = [];
	doc.nodesBetween(from, to, (node, pos) => {
		if (isMarker(node) && pos >= from && pos + node.nodeSize <= to)
			markers.push({ pos, size: node.nodeSize });
	});
	return markers;
}

/**
 * Replaces `from..to` with `text` but keeps any field markers inside the range when removing them
 * would break a field, as Word does when a selection crosses a field boundary.
 */
export function replaceAroundMarkers(
	state: EditorState,
	from: number,
	to: number,
	text: string,
): Transaction | null {
	if (from === to) return null;
	const markers = markersIn(state.doc, from, to);
	if (!markers.length) return null;
	const whole = state.tr.delete(from, to);
	if (fieldsBalanced(whole.doc) || !fieldsBalanced(state.doc)) return null;
	const marks = state.doc.resolve(from).marksAcross(state.doc.resolve(to)) ?? [];
	const transaction = state.tr;
	let end = to;
	for (const marker of [...markers].reverse()) {
		const after = marker.pos + marker.size;
		if (after < end) transaction.delete(after, end);
		end = marker.pos;
	}
	if (from < end) transaction.delete(from, end);
	if (text) transaction.insert(from, state.schema.text(text, marks));
	return transaction.setSelection(TextSelection.create(transaction.doc, from + text.length));
}

/** Backspace/Delete that steps over hidden field markers instead of deleting them. */
export function deleteAroundMarkers(state: EditorState, backward: boolean): Transaction | null {
	const { selection, doc } = state;
	if (!(selection instanceof TextSelection)) return null;
	if (!selection.empty) return replaceAroundMarkers(state, selection.from, selection.to, '');
	let pos = selection.from;
	const step = () => (backward ? doc.resolve(pos).nodeBefore : doc.resolve(pos).nodeAfter);
	if (!isMarker(step())) return null;
	while (isMarker(step())) pos += backward ? -doc.resolve(pos).nodeBefore!.nodeSize : 1;
	const neighbor = step();
	if (!neighbor?.isText) return state.tr.setSelection(TextSelection.create(doc, pos));
	const transaction = backward ? state.tr.delete(pos - 1, pos) : state.tr.delete(pos, pos + 1);
	const caret = backward ? pos - 1 : pos;
	return transaction.setSelection(TextSelection.create(transaction.doc, caret));
}
