// Operating legacy form fields in the editor: a check box toggles and a drop-down picks an entry by
// rewriting the field data that rides on the begin marker; a text field's value is its result text.
import type { Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { FormField, TextRun } from '../model';
import {
	formCheckboxChecked,
	formFieldDisplayText,
	selectFormDropDownEntry,
	setFormCheckbox,
} from '../form-fields';
import { violatesControlLock } from './content-control-guard';
import { inlineRunProperties } from './inline-run-properties';

/** Marks a transaction made by operating a form field, which form protection lets through. */
export const FORM_FIELD_EDIT_META = 'dve-form-field-edit';

/** The form field on a begin marker node, if it is one. */
export function formFieldOfNode(node: Node | null | undefined): FormField | undefined {
	if (node?.type.name !== 'fieldMarker' || node.attrs.kind !== 'begin') return undefined;
	return inlineRunProperties(node).formField;
}

/** What a begin marker node draws for its field: a ballot box, a drop-down's entry, or nothing. */
export function formFieldNodeDisplay(node: Node): string | undefined {
	const formField = formFieldOfNode(node);
	return formField
		? formFieldDisplayText({ text: '', fieldChar: 'begin', formField } as TextRun)
		: undefined;
}

function withField(state: EditorState, pos: number, field: FormField): Transaction {
	const node = state.doc.nodeAt(pos)!;
	const format = typeof node.attrs.format === 'string' ? JSON.parse(node.attrs.format) : {};
	return state.tr
		.setNodeAttribute(pos, 'format', JSON.stringify({ ...format, formField: field }))
		.setMeta(FORM_FIELD_EDIT_META, true);
}

/** Whether a content control around the field at `pos` locks its content, so it cannot change. */
export function formFieldLocked(state: EditorState, pos: number): boolean {
	const node = state.doc.nodeAt(pos);
	return (
		Boolean(node) &&
		violatesControlLock(state.tr.setNodeAttribute(pos, 'format', node!.attrs.format), state)
	);
}

const fillable = (
	state: EditorState,
	pos: number,
	field: FormField | undefined,
): field is FormField => Boolean(field) && field!.enabled !== false && !formFieldLocked(state, pos);

/** Toggles the check box form field whose begin marker is at `pos`. */
export function toggleFormCheckboxAt(state: EditorState, pos: number): Transaction | undefined {
	const field = formFieldOfNode(state.doc.nodeAt(pos));
	if (!fillable(state, pos, field) || field.kind !== 'checkbox') return undefined;
	return withField(state, pos, setFormCheckbox(field, !formCheckboxChecked(field)));
}

/** Selects entry `index` of the drop-down form field whose begin marker is at `pos`. */
export function selectFormDropDownAt(
	state: EditorState,
	pos: number,
	index: number,
): Transaction | undefined {
	const field = formFieldOfNode(state.doc.nodeAt(pos));
	const next =
		fillable(state, pos, field) &&
		field.kind === 'dropDown' &&
		selectFormDropDownEntry(field, index);
	return next ? withField(state, pos, next) : undefined;
}

/** A text form field's result in the editor document: where it is and the field it belongs to. */
export interface FormTextResult {
	field: FormField;
	/** Document position of the field's begin marker. */
	begin: number;
	/** Document positions around the result text (equal when the result is empty). */
	from: number;
	to: number;
	length: number;
}

/** Every text form field's result range, in document order. */
export function formTextResults(doc: Node): FormTextResult[] {
	const results: FormTextResult[] = [];
	// Fields nest, and a result can run on into the next paragraph: track the open begin markers.
	const open: { field: FormField | undefined; begin: number; from?: number; length: number }[] = [];
	doc.descendants((node, pos) => {
		if (!node.isTextblock) return true;
		node.forEach((child, offset) => {
			const at = pos + 1 + offset;
			if (child.type.name !== 'fieldMarker') {
				const top = open.at(-1);
				if (top?.from !== undefined) top.length += child.isText ? child.text!.length : 1;
				return;
			}
			if (child.attrs.kind === 'begin')
				open.push({ field: formFieldOfNode(child), begin: at, length: 0 });
			else if (child.attrs.kind === 'separate') {
				const top = open.at(-1);
				if (top) top.from = at + child.nodeSize;
			} else if (child.attrs.kind === 'end') {
				const top = open.pop();
				if (top?.field?.kind === 'text' && top.from !== undefined)
					results.push({
						field: top.field,
						begin: top.begin,
						from: top.from,
						to: at,
						length: top.length,
					});
			}
		});
		return false;
	});
	return results;
}
