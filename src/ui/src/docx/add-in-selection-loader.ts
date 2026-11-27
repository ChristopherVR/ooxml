import type { EditorView } from 'prosemirror-view';

/** Loads the Word selection bridge; it looks up the current view on each request, not now. */
export async function loadWordAddInSelection(current: () => EditorView | undefined) {
	const { wordAddInSelection } = await import('./add-in-selection');
	return wordAddInSelection(current);
}
