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

export type ReviewRecordingPreferences = Readonly<
	Partial<Record<'trackFormatting' | 'trackMoves', boolean>>
>;

/** Apply changed document preferences together as one isolated history operation. */
export function setReviewRecordingPreferences(preferences: ReviewRecordingPreferences): Command {
	return (state, dispatch, view) => {
		if (view && !view.editable) return false;
		const tr = state.tr;
		for (const name of ['trackFormatting', 'trackMoves'] as const) {
			const value = preferences[name];
			if (value === undefined) continue;
			if (!state.doc.type.spec.attrs?.[name]) return false;
			if ((state.doc.attrs[name] !== false) !== value) tr.setDocAttribute(name, value);
		}
		if (dispatch) dispatchIsolatedCommand(state, dispatch, tr);
		return true;
	};
}
