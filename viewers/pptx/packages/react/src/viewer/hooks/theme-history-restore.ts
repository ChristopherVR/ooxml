import {
	historyTheme,
	loadedThemeName,
	withHistoryMasterThemes,
	withHistoryTheme,
	writeThemeParts,
} from 'ooxml-ui/pptx';
import type { HistoryMaster, HistoryTheme } from 'ooxml-ui/pptx';

import type { UseThemeHandlersInput } from './useThemeHandlers';

/**
 * An undo or redo crossed a theme change: put `target`'s colour scheme, font scheme and name
 * back, in the viewer's theme and in the handler's theme parts, so what a save writes matches the
 * slides again. Only the parts that differ are written (the shared `writeThemeParts`, on the
 * handler's theme write queue): undoing a fonts pick leaves the colour scheme part as it is. A step
 * across a theme applied by path puts the masters back on their recorded theme parts first.
 */
export function restoreHistoryTheme(
	input: UseThemeHandlersInput,
	target: HistoryTheme,
): Promise<void> {
	const masters = input.slideMasters as unknown as HistoryMaster[];
	const current = historyTheme(input.theme, masters);
	input.setTheme((prev) => withHistoryTheme(prev, target));
	if (target.masterThemePaths) {
		input.setSlideMasters(
			(prev) =>
				withHistoryMasterThemes(prev as unknown as HistoryMaster[], target) as unknown as Array<
					Record<string, unknown>
				>,
		);
	}
	return writeThemeParts(
		input.handlerRef.current,
		current,
		target,
		loadedThemeName(input.themeOptions, input.slideMasters),
	);
}
