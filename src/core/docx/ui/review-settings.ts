import type { Command } from 'prosemirror-state';
import { dispatchIsolatedCommand } from './command-history';

/** Document-wide recording state belongs in transactions for sync and undo. */
function toggleRecordingAttribute(
	name: 'trackChanges' | 'trackFormatting' | 'trackMoves',
): Command {
	return (state, dispatch, view) => {
		if ((view && !view.editable) || !state.doc.type.spec.attrs?.[name]) return false;
		if (dispatch)
			dispatchIsolatedCommand(
				state,
				dispatch,
				state.tr.setDocAttribute(name, !state.doc.attrs[name]),
			);
		return true;
	};
}

export const toggleTrackChanges = toggleRecordingAttribute('trackChanges');
export const toggleTrackFormatting = toggleRecordingAttribute('trackFormatting');
export const toggleTrackMoves = toggleRecordingAttribute('trackMoves');
