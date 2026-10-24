import { commitsThroughDeckStyle, galleryDeckView, tableLayoutCells } from 'ooxml-ui/pptx';
import type { RibbonGalleryApplyResult, RibbonGalleryContext } from 'ooxml-ui/pptx';
/**
 * ribbon-gallery-helpers.ts: the Angular glue between `pptx-viewer-shared`'s
 * ribbon galleries and this binding's state services.
 *
 * Shared decides what a gallery offers and what a pick writes
 * (`buildRibbonGallery` / `applyRibbonGalleryItem`); these functions only
 * (1) assemble the {@link RibbonGalleryContext} from the loaded deck and
 * (2) route an apply result onto the binding's existing update paths: the
 * editor's undoable `updateElement` for element patches, and the theme
 * gallery service (the Design > Themes path) for colour / font schemes.
 * Kept free of `@Component` so they are testable without TestBed.
 */
import type { DeckStyleHost } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';

import type { EditorStateService } from './editor-state.service';
import type { ViewerThemeGalleryService } from './viewer-theme-gallery.service';

/** The slice of `LoadContentService` a gallery context reads. */
export interface GalleryDeckSource {
	theme: () => PptxTheme | undefined;
	themeColorMap: () => Record<string, string> | undefined;
	getHandler: () => PptxHandler | undefined;
}

/** The slice of `TableSelectionService.selection()` a gallery context reads. */
export interface GalleryTableSelection {
	elementId: string;
	rowIndex: number;
	columnIndex: number;
	selectedCells?: ReadonlyArray<{ row: number; col: number }> | undefined;
}

/** The context every gallery descriptor is built from. */
export function galleryContextFor(
	element: PptxElement | null,
	deck: GalleryDeckSource | null,
	tableSelection?: GalleryTableSelection | null,
	activeSlide?: PptxSlide | null,
	deckStyle?: DeckStyleHost | null,
): RibbonGalleryContext {
	const handler = deck?.getHandler();
	// While a hover preview shows, galleries are built from the original deck (stable tiles).
	const view = galleryDeckView(deckStyle?.held(), {
		theme: deck?.theme(),
		themeColorMap: deck?.themeColorMap(),
		activeSlide,
	});
	return {
		element,
		themeColorMap: view.themeColorMap,
		theme: view.theme,
		resolveStyleMatrix: handler
			? (xml, effectSchemeId) => handler.resolveStyleMatrixReferences(xml, effectSchemeId)
			: undefined,
		activeSlide: view.activeSlide ?? null,
		...(handler?.getThemeVariants().length ? { themeVariants: handler.getThemeVariants() } : {}),
		// Only a selection on the selected table counts (a stale one from another table is ignored).
		selectedTableCells: tableLayoutCells(
			tableSelection && tableSelection.elementId === element?.id ? tableSelection : null,
		),
	};
}

/** What {@link dispatchGalleryResult} writes through. */
export interface GalleryDispatchTargets {
	editor: Pick<EditorStateService, 'updateElement'>;
	slideIndex: number;
	themes: Pick<ViewerThemeGalleryService, 'applyThemeVariant'> | null;
	/** The Design galleries' hover preview and their effect and background commits. */
	deckStyle?: DeckStyleHost | null;
}

/**
 * Apply a gallery pick. Returns false when nothing could take it (a `null`
 * result, or a theme pick outside a viewer that owns a theme path).
 */
export function dispatchGalleryResult(
	result: RibbonGalleryApplyResult | null,
	targets: GalleryDispatchTargets,
): boolean {
	if (!result) {
		return false;
	}
	if (targets.deckStyle && commitsThroughDeckStyle(result)) {
		// Settles the hover preview itself, so the commit starts from the real slides, writes only
		// what the pick changes into the theme part and records one undo step.
		void targets.deckStyle.commit(result);
		return true;
	}
	if (
		result.kind === 'themeEffects' ||
		result.kind === 'backgroundStyle' ||
		result.kind === 'sequence'
	) {
		// Nothing but the deck-style host commits these.
		return true;
	}
	targets.deckStyle?.endPreview();
	if (result.kind === 'command') {
		return true;
	}
	if (result.kind === 'element') {
		targets.editor.updateElement(targets.slideIndex, result.elementId, result.patch);
		return true;
	}
	if (!targets.themes) {
		return false;
	}
	if (result.kind === 'theme') {
		targets.themes.applyThemeVariant(
			{ colorScheme: result.colorScheme, fontScheme: result.fontScheme },
			result.name,
		);
	} else if (result.kind === 'themeColorScheme') {
		targets.themes.applyThemeVariant({ colorScheme: result.colorScheme }, result.name);
	} else {
		targets.themes.applyThemeVariant({ fontScheme: result.fontScheme }, result.name);
	}
	return true;
}
