import { panelOf, setTabHidden } from './ribbon-tab-api';
import type { EditorState } from 'prosemirror-state';

/** Whether the selection sits inside a table cell. */
export function selectionInTable(state: EditorState): boolean {
	const { $from } = state.selection;
	for (let depth = $from.depth; depth > 0; depth--)
		if ($from.node(depth).type.name === 'table') return true;
	return false;
}

/**
 * Word shows its table tools only while a table is selected. The tab is added to the tab strip
 * when the selection enters a table and removed (falling back to Home) when it leaves.
 */
export function syncContextualTabs(toolbar: HTMLElement, state: EditorState): void {
	const panel = panelOf(toolbar, 'table');
	if (!panel) return;
	setTabHidden(toolbar, 'table', 'contextual', !selectionInTable(state));
}
