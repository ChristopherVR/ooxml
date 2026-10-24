import { applyThemeByPath } from 'ooxml-ui/pptx';
import type { PptxSlideMaster } from 'pptx-viewer-core';

import type { UseThemeHandlersInput } from './useThemeHandlers';

/**
 * Presentation properties > Theme: point the first master, or every one, at another theme part
 * already in the package, without reloading the deck (the slide on screen, the selection and Undo
 * stay). The shared `applyThemeByPath` re-points the masters and re-themes the slides drawn from
 * them; the result lands as one undo step, which records the masters' theme parts, so Undo points
 * them back (`theme-history-restore.ts`).
 */
export async function applyThemePartInPlace(
	input: UseThemeHandlersInput,
	themePath: string,
	applyToAllMasters: boolean,
): Promise<void> {
	const change = await applyThemeByPath(input.handlerRef.current, {
		themePath,
		applyToAllMasters,
		slides: input.slides,
		slideMasters: input.slideMasters as unknown as PptxSlideMaster[],
		theme: input.theme,
		templateElementsBySlideId: input.templateElementsBySlideId,
	});
	if (!change) {
		return;
	}
	input.bumpHistory();
	input.setSlides(change.slides);
	input.setTemplateElementsBySlideId(change.templateElementsBySlideId);
	input.setSlideMasters(change.slideMasters as unknown as Array<Record<string, unknown>>);
	if (change.deckTheme) {
		input.setTheme(change.deckTheme.theme as unknown as Record<string, unknown>);
	}
	input.history.markDirty();
}
