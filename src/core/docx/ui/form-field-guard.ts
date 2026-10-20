// Enforces what legacy form fields allow while editing: a text field's maximum length, and a
// document protected for filling in forms (`w:documentProtection w:edit="forms"`), where only
// form fields and content controls can change.
import type { Node } from 'prosemirror-model';
import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AttrStep,
	RemoveMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import type { DocumentProtection } from '../model';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlSpans } from './content-control-spans';
import { FORM_FIELD_EDIT_META, formFieldOfNode, formTextResults } from './form-field-commands';

const isRemote = (tr: Transaction): boolean => tr.getMeta('y-sync$') !== undefined;

/** Whether the transaction makes a text field's result longer than its maximum length. */
export function exceedsFormTextLength(tr: Transaction): boolean {
	// The same field before and after: its begin marker, mapped through the transaction.
	const before = new Map(
		formTextResults(tr.before).map((result) => [tr.mapping.map(result.begin), result.length]),
	);
	return formTextResults(tr.doc).some((result) => {
		const max = result.field.textInput?.maxLength;
		return Boolean(max) && result.length > max! && result.length > (before.get(result.begin) ?? 0);
	});
}

/** Whether every step of a transaction stays inside a form field or a content control. */
export function staysInFormFields(tr: Transaction): boolean {
	const doc = tr.before;
	const allowed = [
		...formTextResults(doc)
			.filter(({ field }) => field.enabled !== false)
			.map(({ from, to }) => ({ from, to })),
		...controlSpans(doc).map(({ from, to }) => ({ from, to })),
	];
	return tr.steps.every((step, index) => {
		const map = tr.mapping.slice(0, index);
		if (step instanceof AttrStep) {
			const node = tr.docs[index]?.nodeAt(step.pos);
			return Boolean(formFieldOfNode(node));
		}
		// Typing, and the field guard re-marking typed text as the field's result, stay in range.
		if (
			!(
				step instanceof ReplaceStep ||
				step instanceof ReplaceAroundStep ||
				step instanceof AddMarkStep ||
				step instanceof RemoveMarkStep
			)
		)
			return false;
		const { from, to } = step as unknown as { from: number; to: number };
		return allowed.some((range) => from >= map.map(range.from, -1) && to <= map.map(range.to, 1));
	});
}

/** Whether the document has a text field with a maximum length, the only fields the guard limits. */
const hasLimitedField = (doc: Node): boolean =>
	formTextResults(doc).some((result) => result.field.textInput?.maxLength);

/** Whether the transaction brings in a field marker (a paste, an undo, a remote change). */
const insertsFieldMarker = (tr: Transaction): boolean =>
	tr.steps.some((step) => {
		if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) return false;
		let found = false;
		step.slice.content.descendants((node) => {
			if (node.type.name === 'fieldMarker') found = true;
			return !found;
		});
		return found;
	});

const limitedFields = new PluginKey<boolean>('formFieldLimits');

/** Rejects local edits that a field's maximum length or the document's protection forbids. */
export function formFieldGuardPlugin(protection: () => DocumentProtection | undefined): Plugin {
	return new Plugin<boolean>({
		key: limitedFields,
		// Whether any field limits its length, so documents without one skip the scan per edit.
		state: {
			init: (_config, state) => hasLimitedField(state.doc),
			apply: (tr, limited) => (insertsFieldMarker(tr) ? hasLimitedField(tr.doc) : limited),
		},
		filterTransaction: (tr, state) => {
			if (!tr.docChanged || isRemote(tr)) return true;
			const limited = limitedFields.getState(state) || insertsFieldMarker(tr);
			if (limited && exceedsFormTextLength(tr)) return false;
			const forms = protection();
			if (!forms?.enforced || forms.edit !== 'forms') return true;
			return (
				Boolean(tr.getMeta(CONTROL_EDIT_META) || tr.getMeta(FORM_FIELD_EDIT_META)) ||
				staysInFormFields(tr)
			);
		},
	});
}
