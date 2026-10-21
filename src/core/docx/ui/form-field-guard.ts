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
	type Step,
} from 'prosemirror-transform';
import type { DocumentProtection } from '../model';
import { CONTROL_EDIT_META } from './content-control-commands';
import { controlSpans } from './content-control-spans';
import {
	FORM_FIELD_EDIT_META,
	formFieldOfNode,
	formTextResults,
	type FormTextResult,
} from './form-field-commands';
import { isBlankFormText, replacedBlanks } from './form-text-blank';

const isRemote = (tr: Transaction): boolean => tr.getMeta('y-sync$') !== undefined;

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
 * The sections a document protected for forms leaves editable (`w:formProt w:val="0"`), placed as
 * Print Layout and the saved document place them: a section ends after the first block with its
 * end id, the last section runs to the end of the document, and a section whose last block is
 * gone is empty, its blocks falling to the next one (deleting a break merges sections).
 */
function unprotectedSections(doc: Node): Range[] {
	if (typeof doc.attrs.sections !== 'string') return [];
	const sections = JSON.parse(doc.attrs.sections) as {
		endsAtBlockId?: string;
		formProtection?: false;
	}[];
	if (!sections.some((section) => section.formProtection === false)) return [];
	const blocks: { id: unknown; end: number }[] = [];
	doc.forEach((node, offset) => blocks.push({ id: node.attrs.id, end: offset + node.nodeSize }));
	const ranges: Range[] = [];
	let start = 0;
	let next = 0;
	sections.forEach((section, index) => {
		let end = start;
		if (index === sections.length - 1) end = doc.content.size;
		else {
			const found = blocks.findIndex(
				(block, at) => at >= next && block.id === section.endsAtBlockId,
			);
			if (found >= 0) {
				end = blocks[found]!.end;
				next = found + 1;
			}
		}
		// Neighbouring unprotected sections make one editable range.
		const last = ranges.at(-1);
		if (section.formProtection === false && end > start) {
			if (last?.to === start) last.to = end;
			else ranges.push({ from: start, to: end });
		}
		start = end;
	});
	return ranges;
}

const isBlock = (node: Node | null | undefined): node is Node =>
	node?.type.name === 'paragraph' || node?.type.name === 'table';

/**
 * Whether the step only gives a paragraph or table another id, as the editor does to the copy a
 * split makes of a paragraph: bookkeeping, not an edit, as long as no section moves with it.
 */
function renamesBlock(step: Step, before: Node | undefined, after: Node): boolean {
	if (!before) return false;
	if (step instanceof AttrStep) return step.attr === 'id' && isBlock(before.nodeAt(step.pos));
	if (!(step instanceof ReplaceAroundStep)) return false;
	const old = before.nodeAt(step.from);
	const next = after.nodeAt(step.from);
	return Boolean(
		isBlock(old) &&
		next &&
		old.nodeSize === step.to - step.from &&
		old.type === next.type &&
		old.content.eq(next.content) &&
		Mark.sameSet(old.marks, next.marks) &&
		Object.keys({ ...old.attrs, ...next.attrs }).every(
			(key) => key === 'id' || old.attrs[key] === next.attrs[key],
		),
	);
}

/**
 * Whether every step of a transaction stays inside a form field, a content control or a section
 * the protection leaves editable.
 */
export function staysInFormFields(tr: Transaction): boolean {
	const doc = tr.before;
	const fields: Range[] = [
		...formTextResults(doc)
			.filter(({ field }) => field.enabled !== false)
			.map(({ from, to }) => ({ from, to })),
		...controlSpans(doc).map(({ from, to }) => ({ from, to })),
	];
	const sections = unprotectedSections(doc);
	// Disabled text fields outside the unprotected sections, from their begin to their end marker.
	const fixed: Range[] = formTextResults(doc)
		.filter(({ field }) => field.enabled === false)
		.map(({ begin, to }) => ({ from: begin, to: to + 1 }))
		.filter((range) => !sections.some((free) => range.from >= free.from && range.to <= free.to));
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
			// Operating a field sets its marker's data; a node in an editable range may change too.
			if (formFieldOfNode(tr.docs[index]?.nodeAt(step.pos))) return true;
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
		// A disabled field stays as it is, also inside a content control the user may edit. A mark
		// step maps no positions, so its own range is what it changes.
		const touchesFixed = (start: number, end: number) =>
			fixed.some((range) => start < map.map(range.to, -1) && end > map.map(range.from, 1));
		let changesFixed =
			(step instanceof AddMarkStep || step instanceof RemoveMarkStep) && touchesFixed(from, to);
		step.getMap().forEach((start, end) => {
			changesFixed ||= touchesFixed(start, end);
		});
		if (changesFixed) return false;
		if (within(fields, from, to)) return true;
		if (!within(sections, from, to)) return false;
		if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) return true;
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
 * Rejects local edits that a field's maximum length or the document's protection forbids, and
 * removes the en spaces of a blank text field that text was typed into.
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
		filterTransaction: (tr, state) => {
			if (!tr.docChanged || isRemote(tr)) return true;
			const limited = textFields.getState(state)?.limited || insertsFieldMarker(tr);
			if (limited && exceedsFormTextLength(tr)) return false;
			if (!isFormsProtected(protection())) return true;
			return (
				Boolean(tr.getMeta(CONTROL_EDIT_META) || tr.getMeta(FORM_FIELD_EDIT_META)) ||
				staysInFormFields(tr)
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
