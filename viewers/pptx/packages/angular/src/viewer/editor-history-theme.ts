/**
 * editor-history-theme.ts: a theme change as an undo step over the loader's theme signals
 * (`render/theme-history` in `ooxml-ui/pptx`).
 *
 * Every snapshot records the loader's theme (its colour scheme, font scheme and name) and colour
 * map; an undo or redo that crosses a theme change puts them back and writes the parts that differ
 * into the handler's theme parts, so a save after it writes the theme the slides show.
 */
import {
	crossedHistoryTheme,
	historyTheme,
	loadedThemeName,
	withHistoryTheme,
	writeThemeParts,
} from 'ooxml-ui/pptx';
import type { HistoryTheme } from 'ooxml-ui/pptx';

import type { LoadContentService } from './load-content.service';

/** The theme part of an undo snapshot. */
export interface RecordedTheme {
	theme?: HistoryTheme;
	themeColorMap?: Record<string, string>;
}

/** What an undo step records of the loader's theme. */
export function recordLoaderTheme(loader: LoadContentService | null): RecordedTheme {
	const theme = historyTheme(loader?.theme());
	const themeColorMap = loader?.themeColorMap();
	return {
		...(theme ? { theme } : {}),
		...(themeColorMap ? { themeColorMap: { ...themeColorMap } } : {}),
	};
}

/** An undo or redo moved from `current` to `target`: put a crossed theme change back. */
export function restoreLoaderTheme(
	loader: LoadContentService | null,
	current: RecordedTheme,
	target: RecordedTheme,
): void {
	const crossed = crossedHistoryTheme(current.theme, target.theme);
	if (!loader || !crossed) {
		return;
	}
	loader.theme.set(withHistoryTheme(loader.theme(), crossed));
	loader.themeColorMap.set(target.themeColorMap ? { ...target.themeColorMap } : undefined);
	void writeThemeParts(
		loader.getHandler(),
		current.theme,
		crossed,
		loadedThemeName(loader.themeOptions(), loader.slideMasters()[0]?.themePath),
	);
}
