// Enforces content control lock settings, routes typing and deleting into controls, and keeps
// verbatim properties from another page out of pasted controls.
import { isHistoryTransaction } from 'prosemirror-history';
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import { AddMarkStep, RemoveMarkStep, ReplaceAroundStep, ReplaceStep } from 'prosemirror-transform';
import { isContentControlLocked, type ContentControl } from '../content-control-model';
import {
	CONTROL_EDIT_META,
	deleteToPlaceholder,
	pasteInCell,
	typeOverPlaceholder,
} from './content-control-commands';
import { nodeControlOf, otherControlSourcesOf, pastedControlSlice } from './content-control-paste';
import { controlSourcesOf, pastedControlSources } from './content-control-sources';
import type { RowControlSpan } from './content-control-rows';
import { controlOfMark, everyControlSpan, type ControlSpan } from './content-control-spans';
import { removeEditedTemporaryControls } from './content-control-temporary';

/** Remote collaboration transactions are applied as they arrive; the guard only filters local edits. */
const isRemote = (tr: Transaction): boolean => tr.getMeta('y-sync$') !== undefined;

type RangeStep = ReplaceStep | ReplaceAroundStep | AddMarkStep | RemoveMarkStep;
const isRangeStep = (step: unknown): step is RangeStep =>
	step instanceof ReplaceStep ||
	step instanceof ReplaceAroundStep ||
	step instanceof AddMarkStep ||
	step instanceof RemoveMarkStep;

/**
 * Whether two controls are one control: the same id, kind and lock, whatever state they hold (a
 * check box's checked state, a picked entry, a placeholder shown).
 */
function sameControl(left: unknown, right: unknown): boolean {
	const [a, b] = [left ?? null, right ?? null] as (Partial<ContentControl> | null)[];
	if (!a || !b) return a === b;
	return a.id === b.id && a.kind === b.kind && (a.lock ?? null) === (b.lock ?? null);
}

/**
 * Whether `step` only gives the node it covers other attributes, its control still that control,
 * in one step over the whole node (`setNodeMarkup`: a table's properties, a cell's shading, a
 * paragraph's alignment, undoing a command on a cell control): the new node is empty and takes the
 * old one's content, so the node, what it holds and its control stay.
 */
function changesAttributesOnly(step: RangeStep, doc: ProseMirrorNode): boolean {
	if (!(step instanceof ReplaceAroundStep)) return false;
	const { from, to, gapFrom, gapTo, insert, slice } = step;
	const node = doc.nodeAt(from);
	const next = slice.content.firstChild;
	return (
		insert === 1 &&
		gapFrom === from + 1 &&
		gapTo === to - 1 &&
		slice.openStart === 0 &&
		slice.openEnd === 0 &&
		slice.content.childCount === 1 &&
		node !== null &&
		next !== null &&
		!node.isLeaf &&
		next.content.size === 0 &&
		node.nodeSize === to - from &&
		next.type === node.type &&
		sameControl(nodeControlOf(node)?.value, nodeControlOf(next)?.value)
	);
}

/** Whether `mark` is the inline control `control`'s own mark. */
const isMarkOf = (mark: Mark, control: ContentControl): boolean =>
	mark.type.name === 'contentControl' && sameControl(controlOfMark(mark), control);

/**
 * Whether `step`, which covers all of `span`, keeps its control. A mark changes no node, so only
 * an inline control's own mark, taken off or replaced by another control's, removes it; text put in
 * place of an inline control's text keeps it when it carries the control, as undoing or redoing a
 * command on it does.
 */
function keepsControl(step: RangeStep, span: ControlSpan | RowControlSpan): boolean {
	if (step instanceof AddMarkStep)
		return (
			span.level !== 'inline' ||
			step.mark.type.name !== 'contentControl' ||
			isMarkOf(step.mark, span.control)
		);
	if (step instanceof RemoveMarkStep)
		return span.level !== 'inline' || !isMarkOf(step.mark, span.control);
	let kept = false;
	if (span.level === 'inline' && step instanceof ReplaceStep)
		step.slice.content.descendants((node) => {
			kept ||= node.marks.some((mark) => isMarkOf(mark, span.control));
			return !kept;
		});
	return kept;
}

/**
 * Whether a step breaks a lock. `contentLocked` forbids editing what is inside the control (it can
 * still be deleted whole); `sdtLocked` forbids deleting the control (its content stays editable).
 * Undo and redo (`history`) may take a control away whole, as undoing the paste that brought it
 * does; they are held to content locks as any other step is.
 */
export function violatesControlLock(tr: Transaction, state: EditorState, history = false): boolean {
	const locked = everyControlSpan(state.doc).filter((span) => span.control.lock);
	if (!locked.length) return false;
	return tr.steps.some((step, index) => {
		if (!isRangeStep(step)) return false;
		const { from, to } = step as { from: number; to: number };
		const inserts = from === to;
		const before = tr.mapping.slice(0, index);
		const kept = changesAttributesOnly(step, tr.docs[index]!);
		return locked.some((span) => {
			const spanFrom = before.map(span.from, 1);
			const spanTo = before.map(span.to, -1);
			const removesWhole = !inserts && from <= spanFrom && to >= spanTo;
			const keeps = kept || keepsControl(step, span);
			if (isContentControlLocked(span.control, 'control') && removesWhole && !keeps && !history)
				return true;
			// Deleting a control is not editing what it holds; a step over all of it that keeps the
			// control is, formatting included.
			if (!isContentControlLocked(span.control, 'content') || (removesWhole && !keeps))
				return false;
			// Other attributes for one node edit that node alone: the control's own node or one in it.
			if (kept) return from >= spanFrom && from < spanTo;
			// A block or cell control's own boundary tokens are not its content, and every
			// position between them (the ends of its text included) is inside it.
			if (span.level !== 'inline') {
				const innerFrom = spanFrom + 1;
				const innerTo = spanTo - 1;
				return inserts ? from >= innerFrom && from <= innerTo : from < innerTo && to > innerFrom;
			}
			return inserts ? from > spanFrom && from < spanTo : from < spanTo && to > spanFrom;
		});
	});
}

export function contentControlPlugin(): Plugin {
	// The verbatim properties of controls a copy, cut or drag in this editor took: a paste keeps them.
	const copied = new Set<string>();
	return new Plugin({
		filterTransaction: (tr, state) =>
			!tr.docChanged ||
			Boolean(tr.getMeta(CONTROL_EDIT_META)) ||
			isRemote(tr) ||
			!violatesControlLock(tr, state, isHistoryTransaction(tr)),
		appendTransaction: removeEditedTemporaryControls,
		props: {
			transformCopied: (slice) => {
				for (const source of controlSourcesOf(slice)) copied.add(source);
				for (const source of otherControlSourcesOf(slice)) copied.add(source);
				return slice;
			},
			transformPasted: (slice, view, plain) =>
				view
					? pastedControlSlice(
							pastedControlSources(slice, view.state.doc, copied),
							view,
							plain,
							copied,
						)
					: slice,
			handlePaste: (view, _event, slice) => {
				if (!view.editable) return false;
				const result = pasteInCell(view.state, slice);
				if (result === 'blocked') return true;
				if (!result) return false;
				// As ProseMirror marks its own paste.
				view.dispatch(result.scrollIntoView().setMeta('paste', true).setMeta('uiEvent', 'paste'));
				return true;
			},
			handleTextInput: (view, from, to, text) => {
				if (!view.editable) return false;
				const result = typeOverPlaceholder(view.state, from, text, to);
				if (result === 'blocked') return true;
				if (!result) return false;
				view.dispatch(result);
				return true;
			},
			handleKeyDown: (view, event) => {
				if (!view.editable || (event.key !== 'Backspace' && event.key !== 'Delete')) return false;
				const tr = deleteToPlaceholder(
					view.state,
					event.key === 'Backspace' ? 'backward' : 'forward',
				);
				if (!tr) return false;
				view.dispatch(tr);
				return true;
			},
		},
	});
}
