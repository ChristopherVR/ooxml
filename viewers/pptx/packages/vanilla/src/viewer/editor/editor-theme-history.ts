import {
	crossedHistoryTheme,
	historyTheme,
	loadedThemeName,
	withHistoryTheme,
	writeThemeParts,
} from 'ooxml-ui/pptx';
import type { HistoryTheme } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';

/**
 * A theme change as an undo step over the viewer store (`render/theme-history`): every snapshot
 * records the store's colour scheme, font scheme and theme name, and an undo or redo that crosses
 * a theme change puts them back and writes the parts that differ into the handler's theme parts,
 * so a save after it writes the theme the slides show.
 */

function storeTheme(state: ViewerState) {
	return { colorScheme: state.colorScheme, fontScheme: state.fontScheme, name: state.themeName };
}

/** What an undo step records of the store's theme. */
export function storeHistoryTheme(state: ViewerState): HistoryTheme | undefined {
	return historyTheme(storeTheme(state));
}

/** An undo or redo moved from `current` to `target`: put a crossed theme change back. */
export function restoreStoreTheme(
	store: Store<ViewerState>,
	handler: PptxHandler | null,
	current: HistoryTheme | undefined,
	target: HistoryTheme | undefined,
): void {
	const crossed = crossedHistoryTheme(current, target);
	if (!crossed) {
		return;
	}
	const state = store.get();
	const next = withHistoryTheme(storeTheme(state), crossed);
	store.set({ colorScheme: next.colorScheme, fontScheme: next.fontScheme, themeName: next.name });
	void writeThemeParts(
		handler,
		current,
		crossed,
		loadedThemeName(state.themeOptions, state.slideMasters[0]?.themePath),
	);
}
