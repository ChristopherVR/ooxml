// In a document protected for forms, a text form field's value takes text but no formatting, as in
// Word: text pasted into it takes the formatting where it lands, as typed text does, and nothing
// else may bring other formatting into it.
import { Fragment, Mark, Slice, type Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { ReplaceStep } from 'prosemirror-transform';
import type { EditorView } from 'prosemirror-view';
import { formTextResults } from './form-field-commands';

type Range = { from: number; to: number };

/**
 * Whether `from`..`to` lies in the value of an enabled text field outside `free` (the sections the
 * protection leaves editable), its edges included.
 */
export function inProtectedFormText(
	doc: Node,
	from: number,
	to: number,
	free: readonly Range[],
): boolean {
	return formTextResults(doc).some(
		(result) =>
			result.field.enabled !== false &&
			from >= result.from &&
			to <= result.to &&
			!free.some((range) => result.from >= range.from && result.to <= range.to),
	);
}

function inFormatting(fragment: Fragment, marks: readonly Mark[]): Fragment {
	const children: Node[] = [];
	fragment.forEach((node) => {
		if (node.isText) children.push(node.mark(marks));
		else children.push(node.isLeaf ? node : node.copy(inFormatting(node.content, marks)));
	});
	return Fragment.fromArray(children);
}

/**
 * The formatting text takes where it is typed at the selection, as ProseMirror types it: the
 * marks a deletion kept, those across a selection it replaces, else those at the caret, where at
 * a value's start, after the field's separator, they are the value's own.
 */
function landingMarks(state: EditorState): readonly Mark[] {
	const { $from, $to, empty } = state.selection;
	if (state.storedMarks) return state.storedMarks;
	if (!empty) return $from.marksAcross($to) ?? $from.marks();
	const after = $from.nodeAfter;
	return !$from.nodeBefore?.isText && after?.isText ? after.marks : $from.marks();
}

/**
 * Pastes `slice` as text in the formatting where it lands when the selection is in a protected
 * text field's value, as ProseMirror pastes otherwise; false, and nothing done, elsewhere or when
 * the clipboard held nothing the editor reads.
 */
export function pasteIntoFormText(view: EditorView, slice: Slice, free: readonly Range[]): boolean {
	const { state } = view;
	const { from, to } = state.selection;
	if (!slice.size || !inProtectedFormText(state.doc, from, to, free)) return false;
	const text = new Slice(
		inFormatting(slice.content, landingMarks(state)),
		slice.openStart,
		slice.openEnd,
	);
	view.dispatch(
		state.tr
			.replaceSelection(text)
			.scrollIntoView()
			.setMeta('paste', true)
			.setMeta('uiEvent', 'paste'),
	);
	return true;
}

/** Marks other than the field's own, which marks a result's text and is no formatting. */
const formatting = (marks: readonly Mark[]): readonly Mark[] =>
	marks.filter((mark) => mark.type.name !== 'field');

/**
 * Whether `step` writes text with formatting that neither the text it replaces nor the text beside
 * it has, nor `stored` (the marks a deletion kept for the text typed next at the caret, given only
 * for a step there): as text dropped into the value with formatting of its own does. Typed text
 * takes the formatting where it lands.
 */
export function writesFormatting(
	step: ReplaceStep,
	doc: Node,
	stored: readonly Mark[] | null,
): boolean {
	const $from = doc.resolve(step.from);
	const allowed: Mark[] = [...$from.marks(), ...(stored ?? [])];
	const around = [$from.nodeBefore, doc.resolve(step.to).nodeAfter];
	for (const node of around) if (node?.isText) allowed.push(...node.marks);
	doc.nodesBetween(step.from, step.to, (node) => {
		if (node.isText) allowed.push(...node.marks);
	});
	let found = false;
	step.slice.content.descendants((node) => {
		found ||= node.isText && formatting(node.marks).some((mark) => !mark.isInSet(allowed));
		return !found;
	});
	return found;
}

/**
 * Whether `tr`, changing nothing else, sets formatting for the text typed next (stored marks other
 * than those at the caret) while the caret is in a protected text field's value, as Bold with no
 * selection does. A deletion that keeps the marks of what it deleted for the text typed next is no
 * such change.
 */
export function storesFormatting(tr: Transaction, free: () => readonly Range[]): boolean {
	if (tr.docChanged || !tr.storedMarksSet || !tr.storedMarks) return false;
	const { from, to, $from } = tr.selection;
	if (Mark.sameSet(formatting(tr.storedMarks), formatting($from.marks()))) return false;
	return inProtectedFormText(tr.doc, from, to, free());
}
