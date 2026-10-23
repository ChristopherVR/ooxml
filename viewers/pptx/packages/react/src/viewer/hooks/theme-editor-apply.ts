import type { ThemeEditorEdit } from 'ooxml-ui/pptx';
import {
	applyThemeToData,
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
} from 'pptx-viewer-core';
import type { PptxData } from 'pptx-viewer-core';

import { queueThemeWrite } from './theme-history-restore';
import type { UseThemeHandlersInput } from './useThemeHandlers';

/**
 * Commit a staged edit without reloading the presentation or resetting selection. It lands as one
 * undo step, which records the theme too, so Undo puts the theme part back as well
 * (`theme-history-restore.ts`).
 */
export async function applyThemeEditorEdit(
	input: UseThemeHandlersInput,
	{ colorScheme, fontScheme, name }: ThemeEditorEdit,
): Promise<void> {
	const {
		handlerRef,
		theme,
		bumpHistory,
		setSlides,
		setTemplateElementsBySlideId,
		setTheme,
		history,
	} = input;

	const handler = handlerRef.current;
	if (!handler) {
		return;
	}
	await queueThemeWrite(handler, () => handler.applyTheme(colorScheme, fontScheme, name));
	const previousMap = theme?.colorScheme ? buildThemeColorMap(theme.colorScheme) : {};
	bumpHistory();
	setSlides(
		(slides) =>
			applyThemeToData(
				{ slides, theme, themeColorMap: previousMap } as PptxData,
				colorScheme,
				fontScheme,
				name,
			).slides,
	);
	setTemplateElementsBySlideId((prev) =>
		Object.fromEntries(
			Object.entries(prev).map(([id, elements]) => [
				id,
				reResolveElementFonts(
					reResolveElementColors(elements, previousMap, colorScheme),
					theme?.fontScheme,
					fontScheme,
				),
			]),
		),
	);
	setTheme((prev) => ({ ...prev, colorScheme, fontScheme, name }));
	history.markDirty();
}
