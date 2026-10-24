import { commitsThroughDeckStyle, createDeckStyleHost } from 'ooxml-ui/pptx';
import type {
	DeckStyleHandler,
	DeckStyleHost,
	DeckStyleState,
	RibbonGalleryApplyResult,
} from 'ooxml-ui/pptx';
import { buildThemeColorMap } from 'pptx-viewer-core';
import type { PptxElement, PptxHandler } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import { getActiveElements, replaceActiveElements } from './editor-active-elements';
import type { EditorOps } from './editor-operations';

/**
 * The ribbon style galleries' dispatcher. What a pick writes is decided by the
 * shared `applyRibbonGalleryItem`; this only routes the result onto the
 * binding's history-integrated paths: an element patch merges onto the named
 * element (push -> mutate -> commit, the same shape every inspector edit takes),
 * and every Design gallery pick (Themes, Colors, Fonts, Effects, Background
 * Styles) goes through the shared deck-style commit, which writes only what the
 * pick changes into the theme part and records one undo step.
 */
export interface GalleryActions {
	applyRibbonGalleryResult(result: RibbonGalleryApplyResult): void;
	/** Show what a Design gallery tile would do while it is hovered; `null` ends the preview. */
	previewRibbonGalleryResult(result: RibbonGalleryApplyResult | null): void;
	/** Put the slides back (the pointer left the tiles, or the gallery closed). */
	endRibbonGalleryPreview(): void;
	/** The original deck while a preview shows; the ribbon builds its galleries from it. */
	heldDeckStyleState(): DeckStyleState | undefined;
}

export interface GalleryActionsDeps {
	store: Store<ViewerState>;
	ops: Pick<EditorOps, 'pushHistory' | 'commitChange'>;
	getHandler?: () => PptxHandler | null;
}

/** The Design galleries' preview and commit over the viewer store. */
function createStoreDeckStyleHost(
	store: Store<ViewerState>,
	ops: Pick<EditorOps, 'pushHistory' | 'commitChange'>,
	getHandler: () => PptxHandler | null,
): DeckStyleHost {
	const write = (state: DeckStyleState): void => {
		store.set({
			slides: state.slides,
			templateElementsBySlideId: state.templateElementsBySlideId,
			...(state.theme?.colorScheme ? { colorScheme: state.theme.colorScheme } : {}),
			...(state.theme?.fontScheme ? { fontScheme: state.theme.fontScheme } : {}),
			...(state.theme?.formatScheme ? { formatScheme: state.theme.formatScheme } : {}),
			...(state.theme?.name ? { themeName: state.theme.name } : {}),
		});
	};
	return createDeckStyleHost(() => ({
		read: () => {
			const state = store.get();
			return {
				slides: state.slides,
				templateElementsBySlideId: state.templateElementsBySlideId,
				theme: {
					name: state.themeName,
					colorScheme: state.colorScheme,
					fontScheme: state.fontScheme,
					formatScheme: state.formatScheme,
				},
				themeColorMap: state.colorScheme ? buildThemeColorMap(state.colorScheme) : undefined,
			};
		},
		show: write,
		hold: (on) => store.set({ deckStyleHeld: on }),
		handler: () => getHandler() as unknown as DeckStyleHandler | null,
		selectedSlideIds: () => {
			const state = store.get();
			const slide = state.slides[state.currentSlide];
			return slide ? [slide.id] : [];
		},
		apply: (state) => {
			ops.pushHistory();
			write(state);
			ops.commitChange();
		},
	}));
}

export function createGalleryActions(deps: GalleryActionsDeps): GalleryActions {
	const { store, ops } = deps;
	const deckStyle = createStoreDeckStyleHost(store, ops, deps.getHandler ?? (() => null));
	return {
		previewRibbonGalleryResult(result) {
			if (store.get().editable) {
				deckStyle.preview(result);
			}
		},
		endRibbonGalleryPreview: () => deckStyle.endPreview(),
		heldDeckStyleState: () => deckStyle.held(),
		applyRibbonGalleryResult(result) {
			const state = store.get();
			if (!state.editable) {
				return;
			}
			if (commitsThroughDeckStyle(result)) {
				// Settles the hover preview itself, so the commit starts from the real slides. A Colors
				// or Fonts pick writes only that scheme and keeps the theme's name.
				void deckStyle.commit(result);
				return;
			}
			deckStyle.endPreview();
			if (result.kind === 'command') {
				return;
			}
			if (result.kind === 'element') {
				const elements = getActiveElements(state);
				if (!elements.some((element) => element.id === result.elementId)) {
					return;
				}
				ops.pushHistory();
				store.set(
					replaceActiveElements(
						state,
						elements.map((element) =>
							element.id === result.elementId
								? ({ ...element, ...result.patch } as PptxElement)
								: element,
						),
					),
				);
				ops.commitChange();
			}
		},
	};
}
