import type { Command } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';

/** Document-wide recording state belongs in transactions for sync and undo. */
export const toggleTrackChanges: Command = (state, dispatch, view) => {
	if ((view && !view.editable) || !state.doc.type.spec.attrs?.trackChanges) return false;
	if (dispatch)
		dispatch(closeHistory(state.tr).setDocAttribute('trackChanges', !state.doc.attrs.trackChanges));
	return true;
};
