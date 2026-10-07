import { closeHistory } from 'prosemirror-history';
import type { EditorState, Transaction } from 'prosemirror-state';
import { wordYjsPluginKey } from './yjs-collaboration.js';

/** Keep one explicit command separate from surrounding typing in both history modes. */
export function dispatchIsolatedCommand(
	state: EditorState,
	dispatch: (transaction: Transaction) => void,
	transaction: Transaction,
): void {
	if (!transaction.docChanged) return;
	const collaboration = wordYjsPluginKey.getState(state);
	collaboration?.stopCapturing();
	try {
		dispatch(closeHistory(transaction));
	} finally {
		collaboration?.stopCapturing();
	}
}
