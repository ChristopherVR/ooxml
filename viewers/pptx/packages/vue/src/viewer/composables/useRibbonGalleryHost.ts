import { commitsThroughDeckStyle, galleryDeckView, tableLayoutCells } from 'ooxml-ui/pptx';
import type {
	DeckStyleHost,
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	TableCursorState,
} from 'ooxml-ui/pptx';
/**
 * useRibbonGalleryHost: the viewer-side half of the ribbon style galleries.
 *
 * Every gallery decision (what a gallery offers, how a tile looks, what a
 * pick writes) is made by `buildRibbonGallery` / `applyRibbonGalleryItem` in
 * `pptx-viewer-shared`. This composable only supplies what those calls need
 * from the loaded deck (the {@link RibbonGalleryContext}) and routes an apply
 * result onto the viewer's existing, history-aware update paths: element
 * patches through `ops.updateElement` (the same path the inspector's Quick
 * Styles use) and scheme swaps through the Design tab's theme editor.
 *
 * `PowerPointViewer` provides it under {@link RibbonGalleryHostKey}; gallery
 * components read it with {@link useRibbonGalleryHost}, which falls back to an
 * inert host (no selection, no-op dispatch) when mounted standalone.
 */
import type { PptxElement, PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { computed, inject, provide } from 'vue';
import type { ComputedRef, InjectionKey, Ref, ShallowRef } from 'vue';

export interface RibbonGalleryHost {
	editable?: ComputedRef<boolean>;
	/** The context every descriptor is built from; rebuilt on selection / deck change. */
	context: ComputedRef<RibbonGalleryContext>;
	/** Carry out what a tile pick asked for. */
	dispatch: (result: RibbonGalleryApplyResult) => void;
	/** Show what a Design gallery tile would do while it is hovered; `null` ends the preview. */
	preview?: (result: RibbonGalleryApplyResult | null) => void;
	/** Put the slides back (the pointer left the tiles, or the gallery closed). */
	endPreview?: () => void;
}

export const RibbonGalleryHostKey: InjectionKey<RibbonGalleryHost> =
	Symbol('pptxRibbonGalleryHost');

/** Where an apply result lands. */
export interface RibbonGallerySinks {
	/** History-tracked element update (shallow top-level merge). */
	updateElement: (id: string, patch: Partial<PptxElement>) => void;
	/** Deck-wide theme update; unset scheme parts keep the current theme's. */
	applyTheme: (updates: Partial<PptxTheme>) => void;
	/** The Design galleries' hover preview and their effect and background commits. */
	deckStyle?: DeckStyleHost | undefined;
}

/** Route one apply result to its sink. */
export function dispatchRibbonGalleryResult(
	result: RibbonGalleryApplyResult,
	sinks: RibbonGallerySinks,
): void {
	if (sinks.deckStyle && commitsThroughDeckStyle(result)) {
		// Settles the hover preview itself, so the commit starts from the real slides, writes only
		// what the pick changes into the theme part and records one undo step (a Fonts pick keeps
		// the colours and the theme's name).
		void sinks.deckStyle.commit(result);
		return;
	}
	sinks.deckStyle?.endPreview();
	switch (result.kind) {
		case 'element':
			sinks.updateElement(result.elementId, result.patch);
			return;
		case 'theme':
			sinks.applyTheme({
				colorScheme: result.colorScheme,
				fontScheme: result.fontScheme,
				name: result.name,
			});
			return;
		case 'themeColorScheme':
			sinks.applyTheme({ colorScheme: result.colorScheme });
			return;
		case 'themeFontScheme':
			sinks.applyTheme({ fontScheme: result.fontScheme });
	}
}

export interface UseRibbonGalleryHostInput extends RibbonGallerySinks {
	editable: () => boolean;
	/** The primary selected element, or null. */
	selectedElement: () => PptxElement | null;
	theme: Ref<PptxTheme | undefined>;
	themeColorMap: Ref<Record<string, string> | undefined>;
	handler: ShallowRef<PptxHandler | null>;
	/** The table cell cursor / range (Table Layout commands act on it). */
	tableSelection?: () => TableCursorState | null;
	/** The slide being edited, for Background Styles' applied mark. */
	activeSlide?: () => PptxSlide | null;
}

/** Build and provide the gallery host for the viewer's ribbon. */
export function provideRibbonGalleryHost(input: UseRibbonGalleryHostInput): RibbonGalleryHost {
	const context = computed<RibbonGalleryContext>(() => {
		const handler = input.handler.value;
		// While a hover preview shows, galleries are built from the original deck (stable tiles).
		const deck = galleryDeckView(input.deckStyle?.held(), {
			theme: input.theme.value,
			themeColorMap: input.themeColorMap.value,
			activeSlide: input.activeSlide?.() ?? null,
		});
		return {
			element: input.selectedElement(),
			theme: deck.theme,
			themeColorMap: deck.themeColorMap,
			resolveStyleMatrix: handler
				? (xml, effectSchemeId) => handler.resolveStyleMatrixReferences(xml, effectSchemeId)
				: undefined,
			selectedTableCells: tableLayoutCells(input.tableSelection?.()),
			activeSlide: deck.activeSlide ?? null,
			...(handler?.getThemeVariants().length ? { themeVariants: handler.getThemeVariants() } : {}),
		};
	});
	const host: RibbonGalleryHost = {
		editable: computed(input.editable),
		context,
		dispatch: (result) => dispatchRibbonGalleryResult(result, input),
		preview: (result) => {
			if (input.editable()) {
				input.deckStyle?.preview(result);
			}
		},
		endPreview: () => input.deckStyle?.endPreview(),
	};
	provide(RibbonGalleryHostKey, host);
	return host;
}

const INERT_HOST: RibbonGalleryHost = {
	context: computed(() => ({ element: null })),
	dispatch: () => {},
};

/** The provided gallery host, or an inert one outside a viewer. */
export function useRibbonGalleryHost(): RibbonGalleryHost {
	return inject(RibbonGalleryHostKey, INERT_HOST);
}
