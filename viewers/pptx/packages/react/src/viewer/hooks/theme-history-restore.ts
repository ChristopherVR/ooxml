import { historyTheme, loadedThemeName, withHistoryTheme, writeThemeParts } from 'ooxml-ui/pptx';
import type { HistoryTheme } from 'ooxml-ui/pptx';

import type { UseThemeHandlersInput } from './useThemeHandlers';

/**
 * An undo or redo crossed a theme change: put `target`'s colour scheme, font scheme and name
 * back, in the viewer's theme and in the handler's theme parts, so what a save writes matches the
 * slides again. Only the parts that differ are written (the shared `writeThemeParts`, on the
 * handler's theme write queue): undoing a fonts pick leaves the colour scheme part as it is.
 */
export function restoreHistoryTheme(
	input: UseThemeHandlersInput,
	target: HistoryTheme,
): Promise<void> {
	const current = historyTheme(input.theme);
	input.setTheme((prev) => withHistoryTheme(prev, target));
	return writeThemeParts(
		input.handlerRef.current,
		current,
		target,
		loadedThemeName(input.themeOptions, input.slideMasters),
	);
}
