import { applyThemeByPath } from 'ooxml-ui/pptx';
import type { PptxHandler } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import type { EditorOps } from './editor-operations';

/**
 * Presentation properties > Theme: point the first master, or every one, at another theme part
 * already in the package, in place, as one undo step.
 *
 * Re-pointing the masters alone (what the deck action did) left the slides in the old theme's
 * colours and fonts until the file was reopened. The shared `applyThemeByPath` re-points the
 * masters and re-themes the slides drawn from them; the store takes its result after one history
 * entry, whose snapshot records the masters' theme parts (`editor-theme-history.ts`), so Undo
 * points them back.
 */
export async function applyThemePartInPlace(
	store: Store<ViewerState>,
	ops: Pick<EditorOps, 'pushHistory' | 'commitChange'>,
	handler: PptxHandler,
	themePath: string,
	applyToAllMasters: boolean,
): Promise<void> {
	const state = store.get();
	const change = await applyThemeByPath(handler, {
		themePath,
		applyToAllMasters,
		slides: state.slides,
		slideMasters: state.slideMasters,
		theme: {
			colorScheme: state.colorScheme,
			fontScheme: state.fontScheme,
			formatScheme: state.formatScheme,
			name: state.themeName,
		},
		templateElementsBySlideId: state.templateElementsBySlideId,
	});
	if (!change) {
		return;
	}
	ops.pushHistory();
	const theme = change.deckTheme?.theme;
	store.set({
		slides: change.slides,
		templateElementsBySlideId: change.templateElementsBySlideId,
		slideMasters: change.slideMasters,
		...(theme
			? {
					colorScheme: theme.colorScheme,
					fontScheme: theme.fontScheme,
					formatScheme: theme.formatScheme,
					themeName: theme.name,
				}
			: {}),
	});
	ops.commitChange();
}
