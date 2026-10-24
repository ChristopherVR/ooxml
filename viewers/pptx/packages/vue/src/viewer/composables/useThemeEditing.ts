import { historyTheme, queueThemeWrite, writeThemeParts } from 'ooxml-ui/pptx';
import { applyThemeToData, reResolveElementColors, reResolveElementFonts } from 'pptx-viewer-core';
import type {
	PptxData,
	PptxElement,
	PptxSlide,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemePreset,
	PptxHandler,
} from 'pptx-viewer-core';
import { ref } from 'vue';
import type { Ref, ShallowRef } from 'vue';

export interface UseThemeEditingInput {
	handler?: Ref<PptxHandler | null | undefined>;
	slides: Ref<PptxSlide[]>;
	pptxTheme: Ref<PptxTheme | undefined>;
	themeColorMap: Ref<Record<string, string> | undefined>;
	pushHistory: () => void;
	themeGalleryOpen: Ref<boolean>;
	themeEditorOpen: Ref<boolean>;
	/**
	 * Master/layout elements rendered as a separate per-slide overlay layer
	 * (not part of `slide.elements`), re-coloured alongside `slides` so an
	 * applied theme doesn't leave inherited background shapes painted with
	 * the old scheme.
	 */
	templateElementsBySlideId: ShallowRef<Record<string, PptxElement[]>>;
	/** The primary theme part's loaded name; the parsed theme carries none. */
	loadedThemeName?: () => string | undefined;
}

export interface UseThemeEditingResult {
	themeEditorBusy: Ref<boolean>;
	/**
	 * Re-theme the deck as one undo step and write the parts that changed into the handler's theme
	 * parts (resolves once they are written). Without a `name` the theme keeps its own.
	 */
	applyTheme: (
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme | undefined,
		name: string | undefined,
	) => Promise<void>;
	/** A partial theme update (the assistant): unset parts keep the current theme's. */
	applyThemeUpdates: (updates: Partial<PptxTheme>) => Promise<void>;
	applyThemePreset: (preset: PptxThemePreset) => Promise<void>;
	applyThemeEdit: (payload: {
		colorScheme: PptxThemeColorScheme;
		fontScheme: PptxThemeFontScheme;
		name: string;
	}) => Promise<void>;
}

/**
 * useThemeEditing: Design ▸ Browse Themes, Edit theme and the assistant's theme updates.
 * Re-themes the whole deck via core's pure `applyThemeToData` (re-resolves slide colours and fonts
 * against the new schemes), writes the new slides/theme/colour-map back as one undo step, and
 * writes the theme into the handler's theme parts so a save carries it. The inline Design
 * galleries commit through the shared deck-style host instead (`useDeckStyleHost`).
 */
export function useThemeEditing(input: UseThemeEditingInput): UseThemeEditingResult {
	const themeEditorBusy = ref(false);
	const {
		slides,
		pptxTheme,
		themeColorMap,
		pushHistory,
		themeGalleryOpen,
		themeEditorOpen,
		templateElementsBySlideId,
	} = input;

	/** The slides, the template layer, the theme and its colour map, as one undo step. */
	function applyThemeInPlace(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme | undefined,
		name: string | undefined,
	): void {
		pushHistory();
		const previousColorMap = themeColorMap.value ?? {};
		const previousFonts = pptxTheme.value?.fontScheme;
		const result = applyThemeToData(
			{
				slides: slides.value,
				theme: pptxTheme.value,
				themeColorMap: previousColorMap,
			} as unknown as PptxData,
			colorScheme,
			fontScheme,
			name,
		);
		slides.value = result.slides;
		pptxTheme.value = result.theme;
		themeColorMap.value = result.themeColorMap;
		if (Object.keys(templateElementsBySlideId.value).length > 0) {
			const recoloured: Record<string, PptxElement[]> = {};
			for (const [slideId, elements] of Object.entries(templateElementsBySlideId.value)) {
				recoloured[slideId] = reResolveElementFonts(
					reResolveElementColors(elements, previousColorMap, colorScheme),
					previousFonts,
					fontScheme ?? previousFonts ?? {},
				);
			}
			templateElementsBySlideId.value = recoloured;
		}
	}
	function applyTheme(
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme | undefined,
		name: string | undefined,
	): Promise<void> {
		const from = historyTheme(pptxTheme.value);
		applyThemeInPlace(colorScheme, fontScheme, name);
		// Only the parts this changed: a fonts-only update leaves the colour scheme part as authored.
		return writeThemeParts(
			input.handler?.value,
			from,
			historyTheme(pptxTheme.value) ?? {},
			input.loadedThemeName?.(),
		);
	}
	function applyThemeUpdates(updates: Partial<PptxTheme>): Promise<void> {
		const current = pptxTheme.value;
		const colorScheme = updates.colorScheme ?? current?.colorScheme;
		if (!colorScheme) {
			return Promise.resolve();
		}
		return applyTheme(colorScheme, updates.fontScheme ?? current?.fontScheme, updates.name);
	}
	/** Apply a built-in theme preset (Design ▸ Browse Themes). */
	function applyThemePreset(preset: PptxThemePreset): Promise<void> {
		const written = applyTheme(preset.colorScheme, preset.fontScheme, preset.name);
		themeGalleryOpen.value = false;
		return written;
	}
	/** Apply edited theme colours/fonts/name (Design ▸ Edit theme). */
	async function applyThemeEdit(payload: {
		colorScheme: PptxThemeColorScheme;
		fontScheme: PptxThemeFontScheme;
		name: string;
	}): Promise<void> {
		if (themeEditorBusy.value) {
			return;
		}
		themeEditorBusy.value = true;
		try {
			const handler = input.handler?.value;
			if (handler) {
				await queueThemeWrite(handler, () =>
					handler.applyTheme(payload.colorScheme, payload.fontScheme, payload.name),
				);
			}
			applyThemeInPlace(payload.colorScheme, payload.fontScheme, payload.name);
			themeEditorOpen.value = false;
		} finally {
			themeEditorBusy.value = false;
		}
	}

	return { applyTheme, applyThemeUpdates, applyThemePreset, applyThemeEdit, themeEditorBusy };
}
