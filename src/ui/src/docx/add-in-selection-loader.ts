import type { EditorView } from 'prosemirror-view';

export async function loadWordAddInSelection(current: () => EditorView | undefined) {
	const view = current();
	if (!view) return;
	const { wordAddInSelection } = await import('./add-in-selection');
	if (current() === view && !view.isDestroyed) return wordAddInSelection(current);
}
