import { queueThemeWrite } from 'ooxml-ui/pptx';
/**
 * The binding's whole-theme apply path, shared by Design > Browse Themes and
 * the theme editor's "Apply to Presentation" (`DeckThemeEditor.svelte`): write
 * the preset into the archive, then re-resolve the live slides in place as one
 * undoable `commitSlides`. The inline Design galleries (Themes, Colors, Fonts)
 * commit through the shared deck-style host instead (`editor-deck-style.ts`).
 */
import type {
	PptxData,
	PptxHandler,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemePreset,
} from 'pptx-viewer-core';
import {
	applyThemeToData,
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
} from 'pptx-viewer-core';

import type { EditorState } from './editor-state.svelte';

/** Recolour the master/layout layer (not part of `slide.elements`) against `colorScheme`. */
function recolourTemplateLayer(
	editor: EditorState,
	previousMap: Record<string, string>,
	colorScheme: PptxThemeColorScheme,
): void {
	if (Object.keys(editor.templateElementsBySlideId).length === 0) {
		return;
	}
	const recoloured: typeof editor.templateElementsBySlideId = {};
	for (const [slideId, elements] of Object.entries(editor.templateElementsBySlideId)) {
		recoloured[slideId] = reResolveElementColors(elements, previousMap, colorScheme);
	}
	editor.templateElementsBySlideId = recoloured;
}

/**
 * Design > Browse Themes: re-theme the PRESENTATION with a gallery preset.
 * Writes the preset into the archive (so it saves), re-resolves every live
 * slide's scheme colours in one undoable `commitSlides`, and returns the next
 * theme for the host to publish.
 */
export async function applyThemePreset(
	editor: EditorState,
	handler: PptxHandler,
	current: PptxTheme | undefined,
	preset: PptxThemePreset,
): Promise<PptxTheme> {
	const previousMap = current?.colorScheme ? buildThemeColorMap(current.colorScheme) : {};
	// On the handler's theme write queue: an undo still writing a theme part, or a save, waits for
	// it rather than interleaving with it.
	await queueThemeWrite(handler, () =>
		handler.applyTheme(preset.colorScheme, preset.fontScheme, preset.name),
	);
	const data = { slides: editor.slides, theme: current, themeColorMap: previousMap } as PptxData;
	const result = applyThemeToData(data, preset.colorScheme, preset.fontScheme, preset.name);
	editor.commitSlides(result.slides);
	recolourTemplateLayer(editor, previousMap, preset.colorScheme);
	editor.templateElementsBySlideId = Object.fromEntries(
		Object.entries(editor.templateElementsBySlideId).map(([id, elements]) => [
			id,
			reResolveElementFonts(elements, current?.fontScheme, preset.fontScheme),
		]),
	);
	return (
		result.theme ?? {
			...current,
			name: preset.name,
			colorScheme: preset.colorScheme,
			fontScheme: preset.fontScheme,
		}
	);
}
