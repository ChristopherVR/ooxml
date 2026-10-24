import { historyTheme, loadedThemeName, restoreCrossedTheme } from 'ooxml-ui/pptx';
import type { HistoryTheme } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';

/**
 * A theme change as an undo step over the viewer store (`render/theme-history`): every snapshot
 * records the store's colour scheme, font scheme and theme name, and each slide master's theme part,
 * and an undo or redo that crosses a theme change puts them back and writes the parts that differ
 * into the handler's theme parts, so a save after it writes the theme the slides show. A step
 * across a theme applied by path points the masters back (the store's masters come back with the
 * snapshot).
 */

function storeTheme(state: ViewerState) {
	return { colorScheme: state.colorScheme, fontScheme: state.fontScheme, name: state.themeName };
}

/** What an undo step records of the store's theme. */
export function storeHistoryTheme(state: ViewerState): HistoryTheme {
	return historyTheme(storeTheme(state), state.slideMasters);
}

/** An undo or redo moved from `current` to `target`: put a crossed theme change back. */
export function restoreStoreTheme(
	store: Store<ViewerState>,
	handler: PptxHandler | null,
	current: HistoryTheme | undefined,
	target: HistoryTheme | undefined,
): void {
	const state = store.get();
	const next = restoreCrossedTheme(
		storeTheme(state),
		current,
		target,
		handler,
		loadedThemeName(state.themeOptions, state.slideMasters),
	);
	if (next) {
		store.set({ colorScheme: next.colorScheme, fontScheme: next.fontScheme, themeName: next.name });
	}
}
