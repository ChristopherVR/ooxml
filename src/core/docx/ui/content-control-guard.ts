// Enforces content control lock settings, routes typing and deleting into controls, and keeps
// verbatim properties from another page out of pasted controls.
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AttrStep,
	RemoveMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import { isContentControlLocked } from '../content-control-model';
import { inlineRunProperties } from './inline-run-properties';
import {
	CONTROL_EDIT_META,
	deleteToPlaceholder,
	typeOverPlaceholder,
} from './content-control-commands';
import { otherControlSourcesOf, pastedControlSlice } from './content-control-paste';
import { rowControlSpans } from './content-control-rows';
import { controlSourcesOf, pastedControlSources } from './content-control-sources';
import { controlSpans } from './content-control-spans';
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
 * Whether a step breaks a lock. `contentLocked` forbids editing what is inside the control (it can
 * still be deleted whole); `sdtLocked` forbids deleting the control (its content stays editable).
 * Changing a node's attributes (a form field's data, the formatting of a picture) edits that node.
 */
export function violatesControlLock(tr: Transaction, state: EditorState): boolean {
	// Row controls are kept in the rows' properties, apart from `controlSpans`.
	const locked = [...controlSpans(state.doc), ...rowControlSpans(state.doc)].filter(
		(span) => span.control.lock,
	);
	return tr.steps.some((step, index) => {
		const before = tr.mapping.slice(0, index);
		if (step instanceof AttrStep) {
			// An inline atom keeps its control with its run properties rather than as a mark.
			const node = tr.docs[index]?.nodeAt(step.pos);
			if (node && isContentControlLocked(inlineRunProperties(node).contentControl, 'content'))
				return true;
			return locked.some(
				(span) =>
					isContentControlLocked(span.control, 'content') &&
					step.pos >= before.map(span.from, 1) &&
					step.pos < before.map(span.to, -1),
			);
		}
		if (!locked.length || !isRangeStep(step)) return false;
		const { from, to } = step as { from: number; to: number };
		const inserts = from === to;
		return locked.some((span) => {
			const spanFrom = before.map(span.from, 1);
			const spanTo = before.map(span.to, -1);
			const removesWhole = !inserts && from <= spanFrom && to >= spanTo;
			if (isContentControlLocked(span.control, 'control') && removesWhole) return true;
			if (!isContentControlLocked(span.control, 'content') || removesWhole) return false;
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
			!violatesControlLock(tr, state),
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
