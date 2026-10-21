// Enforces what legacy form fields allow while editing: a text field's maximum length, and a
// document protected for filling in forms (`w:documentProtection w:edit="forms"`), where only
// form fields, content controls and sections left unprotected can change.
import { Mark, type Node } from 'prosemirror-model';
import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AttrStep,
	Mapping,
	RemoveMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import type { DocumentProtection } from '../model';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlSpans } from './content-control-spans';
import { FIELD_RESULT_REPAIR_META } from './field-guard';
import {
	FORM_FIELD_EDIT_META,
	formTextResults,
	operatesFormField,
	type FormTextResult,
} from './form-field-commands';
import { isBlankFormText, replacedBlanks } from './form-text-blank';
import { keepsSections, renamesBlock, unprotectedSections } from './form-protection-sections';
import { pasteIntoFormText, storesFormatting, writesFormatting } from './form-text-paste';
import { trackChangesPluginKey } from './track-changes-mode';

const isRemote = (tr: Transaction): boolean => tr.getMeta('y-sync$') !== undefined;
/** An undo or redo, which puts back text the field held before, whatever its length. */
const isHistory = (tr: Transaction): boolean => tr.getMeta('history$') !== undefined;

/**
 * Whether the document is protected for filling in forms: only its form fields, content controls
 * and unprotected sections can change, and its headers, footers and notes not at all.
 */
export const isFormsProtected = (protection: DocumentProtection | undefined): boolean =>
	Boolean(protection?.enforced) && protection?.edit === 'forms';

/**
 * Whether the transaction makes a text field's result longer than its maximum length. Word's blank
 * result (en spaces only) counts as empty, and so do the en spaces of a blank that text is typed
 * into, which the guard then removes.
 */
export function exceedsFormTextLength(tr: Transaction): boolean {
	const beforeResults = formTextResults(tr.before);
	const afterResults = formTextResults(tr.doc);
	const replaced = replacedBlanks(tr.before, beforeResults, tr.doc, afterResults, tr.mapping);
	const length = (doc: Node, result: FormTextResult, removed: readonly Range[] = []) =>
		isBlankFormText(doc, result)
			? 0
			: result.length -
				removed.filter((space) => space.from >= result.from && space.to <= result.to).length;
	// The same field before and after: its begin marker, mapped through the transaction.
	const before = new Map(
		beforeResults.map((result) => [tr.mapping.map(result.begin), length(tr.before, result)]),
	);
	return afterResults.some((result) => {
		const max = result.field.textInput?.maxLength;
		const now = length(tr.doc, result, replaced);
		return Boolean(max) && now > max! && now > (before.get(result.begin) ?? 0);
	});
}

type Range = { from: number; to: number };

/**
 * Whether every step of a transaction stays inside a form field, a content control or a section
 * the protection leaves editable. `kept` are the marks the state keeps for the text typed next at
 * its caret, `at`.
 */
export function staysInFormFields(
	tr: Transaction,
	kept?: { marks: readonly Mark[]; at: number } | null,
): boolean {
	const doc = tr.before;
	const results = formTextResults(doc);
	const values: Range[] = results.filter(({ field }) => field.enabled !== false);
	const fields: Range[] = [...values, ...controlSpans(doc).map(({ from, to }) => ({ from, to }))];
	const sections = unprotectedSections(doc);
	const free = (range: Range) => sections.some((it) => range.from >= it.from && range.to <= it.to);
	// Disabled text fields outside the unprotected sections, from their begin to their end marker.
	const fixed: Range[] = results
		.filter(({ field }) => field.enabled === false)
		.map(({ begin, to }) => ({ from: begin, to: to + 1 }))
		.filter((range) => !free(range));
	// Values that take text, not formatting, as in Word: a mark step there must be the editor's own
	// (marking typed text as the result, a tracked change), or an undo or redo.
	const plain = values.filter((range) => !free(range));
	const internal = isHistory(tr) || Boolean(tr.getMeta(FIELD_RESULT_REPAIR_META));
	const tracked = Boolean(tr.getMeta(trackChangesPluginKey));
	let sectionsAfter: Range[] | undefined;
	return tr.steps.every((step, index) => {
		const after = tr.docs[index + 1] ?? tr.doc;
		// Sections end at blocks named by id: a new id must leave the unprotected ones as they were.
		if (
			renamesBlock(step, tr.docs[index], after) &&
			JSON.stringify(unprotectedSections(tr.docs[index]!)) ===
				JSON.stringify(unprotectedSections(after))
		)
			return true;
		const map = tr.mapping.slice(0, index);
		const within = (ranges: Range[], from: number, to: number) =>
			ranges.some((range) => from >= map.map(range.from, -1) && to <= map.map(range.to, 1));
		if (step instanceof AttrStep) {
			// Operating a field sets its marker's data (its own commands pass with their meta; this is
			// their undo and redo), but not the marker's formatting, as in Word, where a protected form
			// takes no formatting. A node in an editable range may change.
			if (operatesFormField(step, tr.docs[index])) return true;
			return within([...fields, ...sections], step.pos, step.pos + 1);
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
		// A disabled field stays as it is, also inside a content control the user may edit, and an
		// enabled text field's value takes no formatting. A mark step maps no positions, so its own
		// range is what it changes.
		const touches = (ranges: Range[], start: number, end: number) =>
			ranges.some((range) => start < map.map(range.to, -1) && end > map.map(range.from, 1));
		const marks = step instanceof AddMarkStep || step instanceof RemoveMarkStep;
		if (!internal && !tracked && (within(plain, from, to) || touches(plain, from, to))) {
			if (marks) return false;
			const stored = kept && from === map.map(kept.at) ? kept.marks : null;
			if (step instanceof ReplaceStep && writesFormatting(step, tr.docs[index]!, stored))
				return false;
		}
		let changesFixed = marks && touches(fixed, from, to);
		step.getMap().forEach((start, end) => {
			changesFixed ||= touches(fixed, start, end);
		});
		if (changesFixed) return false;
		// A field or control may change wherever it is, but not move blocks between sections:
		// Enter in a field that ends a protected section would give the rest of its paragraph to
		// the next section.
		if (within(fields, from, to))
			return !sections.length || keepsSections(step, tr.docs[index]!, after);
		if (!within(sections, from, to)) return false;
		if (marks) return true;
		// What the step wrote has to stay in an unprotected section as well: splitting the paragraph
		// that ends one moves the text after the split into the next section.
		const rest = tr.mapping.slice(index + 1);
		const own = step.getMap();
		const start = rest.map(own.map(from, -1), -1);
		const end = rest.map(own.map(to, 1), 1);
		sectionsAfter ??= unprotectedSections(tr.doc);
		return sectionsAfter.some((range) => start >= range.from && end <= range.to);
	});
}

/**
 * Whether the document has text fields, whose blank results typing replaces, and one with a
 * maximum length, the only fields the guard limits.
 */
function textFieldsOf(doc: Node): FieldState {
	const results = formTextResults(doc);
	return {
		text: results.length > 0,
		limited: results.some((result) => result.field.textInput?.maxLength),
	};
}

interface FieldState {
	text: boolean;
	limited: boolean;
}

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

const textFields = new PluginKey<FieldState>('formFieldLimits');

/**
 * Rejects local edits that a field's maximum length or the document's protection forbids, removes
 * the en spaces of a blank text field that text was typed into, and pastes into a protected text
 * field's value as text in the formatting where it lands.
 */
export function formFieldGuardPlugin(protection: () => DocumentProtection | undefined): Plugin {
	return new Plugin<FieldState>({
		key: textFields,
		// Whether there are text fields, and whether one limits its length, so documents without
		// them skip the scans per edit.
		state: {
			init: (_config, state) => textFieldsOf(state.doc),
			apply: (tr, fields) => (insertsFieldMarker(tr) ? textFieldsOf(tr.doc) : fields),
		},
		props: {
			handlePaste: (view, _event, slice) =>
				isFormsProtected(protection()) &&
				pasteIntoFormText(view, slice, unprotectedSections(view.state.doc)),
		},
		filterTransaction: (tr, state) => {
			// A caret in a protected text field's value takes no formatting for the text typed next.
			if (isFormsProtected(protection()) && storesFormatting(tr, () => unprotectedSections(tr.doc)))
				return false;
			if (!tr.docChanged || isRemote(tr)) return true;
			const limited = textFields.getState(state)?.limited || insertsFieldMarker(tr);
			if (limited && !isHistory(tr) && exceedsFormTextLength(tr)) return false;
			if (!isFormsProtected(protection())) return true;
			return (
				Boolean(tr.getMeta(CONTROL_EDIT_META) || tr.getMeta(FORM_FIELD_EDIT_META)) ||
				staysInFormFields(
					tr,
					state.storedMarks && { marks: state.storedMarks, at: state.selection.from },
				)
			);
		},
		appendTransaction: (transactions, old, state) => {
			// A remote change comes with the other side's own removal.
			if (!transactions.some((tr) => tr.docChanged) || transactions.some(isRemote)) return null;
			if (!textFields.getState(old)?.text) return null;
			const mapping = new Mapping();
			for (const tr of transactions) mapping.appendMapping(tr.mapping);
			const replaced = replacedBlanks(
				old.doc,
				formTextResults(old.doc),
				state.doc,
				formTextResults(state.doc),
				mapping,
			);
			if (!replaced.length) return null;
			const tr = state.tr;
			for (const space of [...replaced].reverse()) tr.delete(space.from, space.to);
			return tr;
		},
	});
}
