/**
 * theme-by-path-commit.ts: Presentation properties > Theme (point the first master, or every one,
 * at another theme part already in the package), in place, as one undo step.
 *
 * Re-pointing the masters alone (what the loader's switch did) left the slides in the old theme's
 * colours and fonts until the file was reopened. The shared `applyThemeByPath` (`ooxml-ui/pptx`)
 * re-points the masters and re-themes the slides drawn from them; this commits its result to the
 * editor and the loader. The undo snapshot records the masters' theme parts
 * (`editor-history-theme.ts`), so Undo points them back.
 */
import { applyThemeByPath } from 'ooxml-ui/pptx';

import type { EditorStateService } from './editor-state.service';
import type { LoadContentService } from './load-content.service';

/** Apply `themePath` to the first master or every master; resolves once committed. */
export async function applyThemePartInPlace(
	editor: EditorStateService,
	loader: LoadContentService,
	themePath: string,
	applyToAllMasters: boolean,
	label: string,
): Promise<void> {
	const change = await applyThemeByPath(loader.getHandler(), {
		themePath,
		applyToAllMasters,
		slides: editor.slides(),
		slideMasters: loader.slideMasters(),
		theme: loader.theme(),
		templateElementsBySlideId: editor.templateElementsBySlideId(),
	});
	if (!change) {
		return;
	}
	// Records the undo snapshot (with the loader's theme and masters as they were) first.
	editor.applyReplacement(change.slides, label);
	editor.templateElementsBySlideId.set(change.templateElementsBySlideId);
	loader.slideMasters.set(change.slideMasters);
	if (change.deckTheme) {
		loader.theme.set(change.deckTheme.theme);
		loader.themeColorMap.set(change.deckTheme.themeColorMap);
	}
}
