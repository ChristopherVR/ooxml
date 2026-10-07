import type { Command } from 'prosemirror-state';
import { canonicalHardBreak } from './hard-break-revisions';

/** Inserts a line break with the selection's run properties and stable collaboration metadata. */
export const insertWordHardBreak: Command = (state, dispatch, view) => {
	const type = state.schema.nodes.hardBreak;
	if (!type || (view && !view.editable)) return false;
	if (!dispatch) return true;
	const { selection } = state;
	const marks =
		state.storedMarks ??
		(selection.empty
			? selection.$from.marks()
			: (selection.$from.marksAcross(selection.$to) ?? []));
	const node = canonicalHardBreak(type.create(null, undefined, marks));
	dispatch(state.tr.replaceSelectionWith(node, false).scrollIntoView());
	return true;
};
