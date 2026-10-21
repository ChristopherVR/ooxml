// Enforces content control lock settings, routes typing and deleting into controls, and keeps
// verbatim properties from another page out of pasted controls.
import { isHistoryTransaction } from 'prosemirror-history';
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AttrStep,
	RemoveMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import { isContentControlLocked, type ContentControl } from '../content-control-model';
import { inlineRunProperties } from './inline-run-properties';
import {
	CONTROL_EDIT_META,
	deleteToPlaceholder,
	pasteInCell,
	typeOverPlaceholder,
} from './content-control-commands';
import { controlOccurrences, type Occurrence } from './content-control-occurrences';
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

/** Whether `node` holds the control of `span`: an inline one as its mark, another as its node's. */
function holdsControl(node: ProseMirrorNode, span: ControlSpan | RowControlSpan): boolean {
	if (span.level === 'inline') return node.marks.some((mark) => isMarkOf(mark, span.control));
	const held = nodeControlOf(node);
	return held?.level === span.level && sameControl(held.value, span.control);
}

/**
 * Whether `step`, which covers all of `span`, keeps its control. A mark changes no node, so only
 * an inline control's own mark, taken off or replaced by another control's, removes it; content put
 * in place of the control's keeps it when it holds the control: text carrying an inline control's
 * mark, as undoing or redoing a command on it does, or a paragraph, table, row or cell holding a
 * block, row or cell control, as a table rebuilt with another row or sorted paragraphs do.
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
	if (step instanceof ReplaceStep)
		step.slice.content.descendants((node) => {
			kept ||= holdsControl(node, span);
			return !kept;
		});
	return kept;
}

/**
 * The locked controls of `doc`, each with the range it holds. A block, row or cell control around
 * several nodes (paragraphs and a table between them, rows, cells) is one, from its first to its
 * last: deleting one of them leaves the control and edits its content.
 */
function lockedControls(doc: ProseMirrorNode): Occurrence[] {
	const locked = everyControlSpan(doc).filter((span) => span.control.lock);
	const inline = locked.filter((span) => span.level === 'inline');
	return [
		...inline.map((span) => ({ key: '', spans: [span], from: span.from, to: span.to })),
		...controlOccurrences(
			doc,
			locked.filter((span) => span.level !== 'inline'),
			[],
		),
	];
}

/**
 * Whether a step breaks a lock. `contentLocked` forbids editing what is inside the control (it can
 * still be deleted whole); `sdtLocked` forbids deleting the control (its content stays editable).
 * Changing a node's attributes (a form field's data, the formatting of a picture) edits that node.
 * Undo and redo (`history`) may take a control away whole, as undoing the paste that brought it
 * does; they are held to content locks as any other step is.
 */
export function violatesControlLock(tr: Transaction, state: EditorState, history = false): boolean {
	const locked = lockedControls(state.doc);
	return tr.steps.some((step, index) => {
		const before = tr.mapping.slice(0, index);
		if (step instanceof AttrStep) {
			// An inline atom keeps its control with its run properties rather than as a mark.
			const node = tr.docs[index]?.nodeAt(step.pos);
			if (node && isContentControlLocked(inlineRunProperties(node).contentControl, 'content'))
				return true;
			return locked.some(
				(occurrence) =>
					isContentControlLocked(occurrence.spans[0]!.control, 'content') &&
					step.pos >= before.map(occurrence.from, 1) &&
					step.pos < before.map(occurrence.to, -1),
			);
		}
		if (!locked.length || !isRangeStep(step)) return false;
		const { from, to } = step as { from: number; to: number };
		const inserts = from === to;
		const kept = changesAttributesOnly(step, tr.docs[index]!);
		return locked.some((occurrence) => {
			const span = occurrence.spans[0]!;
			const spanFrom = before.map(occurrence.from, 1);
			const spanTo = before.map(occurrence.to, -1);
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
