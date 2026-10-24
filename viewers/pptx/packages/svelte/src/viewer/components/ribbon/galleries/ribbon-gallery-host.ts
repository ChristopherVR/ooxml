import {
	applyRibbonGalleryItem,
	buildRibbonGallery,
	commitsThroughDeckStyle,
	galleryDeckView,
} from 'ooxml-ui/pptx';
import type {
	DeckStyleHost,
	RibbonGalleryApplyResult,
	RibbonGalleryContext,
	RibbonGalleryDescriptor,
	RibbonGalleryId,
} from 'ooxml-ui/pptx';
/**
 * The ribbon galleries' link to the editor: it builds the shared
 * `RibbonGalleryContext` from the editor's live selection and deck, asks
 * shared for descriptors, and dispatches a pick's `RibbonGalleryApplyResult`
 * onto this binding's existing update paths (the undoable element patch the
 * inspector uses, and the shared Design gallery commit for every theme, scheme,
 * effect and background pick). Design > Browse Themes and Edit Theme reach the
 * presentation theme through the same host. Every decision about
 * WHAT a gallery offers lives in `pptx-viewer-shared`.
 *
 * Published to the ribbon subtree through Svelte context so every tab (Home,
 * Design and the contextual tabs) reaches the same host without prop
 * threading.
 */
import type { PptxTheme, PptxThemePreset, XmlObject } from 'pptx-viewer-core';
import { getContext, setContext } from 'svelte';

import { createEditorDeckStyleHost } from '../../../editor/editor-deck-style';
import type { EditorState } from '../../../editor/editor-state.svelte';
import { applyThemePreset } from '../../../editor/editor-theme-scheme';

export interface RibbonGalleryHost {
	/** The context shared galleries read; reading it inside `$derived` tracks the editor. */
	context(): RibbonGalleryContext;
	build(id: RibbonGalleryId): RibbonGalleryDescriptor;
	/**
	 * Apply tile `itemId` of gallery `id` to the selection / deck; `action` is the tile's context
	 * menu entry. Resolves the result when it is a command for the caller to run.
	 */
	apply(
		id: RibbonGalleryId,
		itemId: string,
		action?: string,
	): Promise<RibbonGalleryApplyResult | undefined>;
	/** Show what tile `itemId` would do while it is hovered; `endPreview` puts the slides back. */
	preview(id: RibbonGalleryId, itemId: string): void;
	endPreview(): void;
	/** Design > Browse Themes: re-theme the presentation with `preset` (undoable). */
	applyThemePreset(preset: PptxThemePreset): Promise<void>;
	/** Publish a new presentation theme (Design > Edit Theme's editor calls this). */
	publishTheme(theme: PptxTheme): void;
}

/** Build the host over `editor`; `onthemechange` publishes a new presentation theme. */
export function createRibbonGalleryHost(
	editor: EditorState,
	onthemechange?: (theme: PptxTheme) => void,
): RibbonGalleryHost {
	const deckStyle: DeckStyleHost = createEditorDeckStyleHost(editor, (theme) => {
		if (theme) {
			onthemechange?.(theme);
		}
	});
	const context = (): RibbonGalleryContext => {
		const handler = editor.getHandler();
		// While a hover preview shows, galleries are built from the original deck (stable tiles).
		const deck = galleryDeckView(deckStyle.held(), {
			theme: editor.theme,
			themeColorMap: editor.themeColorMap,
			activeSlide: editor.slides[editor.currentSlideIndex] ?? null,
		});
		return {
			element: editor.selectedElement ?? null,
			themeColorMap: deck.themeColorMap,
			theme: deck.theme,
			resolveStyleMatrix: handler
				? (xml: XmlObject, effectSchemeId?: string) =>
						handler.resolveStyleMatrixReferences(xml, effectSchemeId)
				: undefined,
			activeSlide: deck.activeSlide ?? null,
			...(handler?.getThemeVariants().length ? { themeVariants: handler.getThemeVariants() } : {}),
			// The marquee'd range (a plain click is a one-cell range); Table Layout acts on it.
			selectedTableCells: editor.tableCells
				.cellsFor(editor.selectedElement?.id)
				.map(({ row, col }) => ({ row, column: col })),
		};
	};

	async function dispatch(result: RibbonGalleryApplyResult): Promise<void> {
		if (commitsThroughDeckStyle(result)) {
			// Settles the hover preview itself, so the commit starts from the real slides. A Themes,
			// Colors or Fonts pick re-resolves the slides, writes only what it changes into the theme
			// part and records one undo step; the host publishes its theme (`onthemechange`).
			await deckStyle.commit(result);
			return;
		}
		deckStyle.endPreview();
		if (result.kind === 'element') {
			editor.applyElementPatch(result.elementId, result.patch);
		}
	}

	function publishTheme(next: PptxTheme): void {
		editor.theme = next;
		onthemechange?.(next);
	}

	return {
		context,
		build: (id) => {
			const descriptor = buildRibbonGallery(id, context());
			return { ...descriptor, disabled: descriptor.disabled || !editor.editable };
		},
		apply: async (id, itemId, action) => {
			if (!editor.editable) {
				return undefined;
			}
			const result = applyRibbonGalleryItem(id, itemId, context(), action);
			if (!result) {
				return undefined;
			}
			await dispatch(result);
			return result.kind === 'command' ? result : undefined;
		},
		preview: (id, itemId) => {
			if (editor.editable) {
				deckStyle.preview(applyRibbonGalleryItem(id, itemId, context()));
			}
		},
		endPreview: () => deckStyle.endPreview(),
		applyThemePreset: async (preset) => {
			const handler = editor.getHandler();
			if (!handler || !editor.editable) {
				return;
			}
			publishTheme(await applyThemePreset(editor, handler, editor.theme, preset));
		},
		publishTheme,
	};
}

const RIBBON_GALLERY_HOST_KEY = Symbol('pptx-svelte-ribbon-gallery-host');

/** Publish `host` to the ribbon subtree (Ribbon.svelte, or a test harness). */
export function provideRibbonGalleryHost(host: RibbonGalleryHost): void {
	setContext(RIBBON_GALLERY_HOST_KEY, host);
}

/** The nearest host, or undefined when a gallery is mounted outside a ribbon. */
export function useRibbonGalleryHost(): RibbonGalleryHost | undefined {
	return getContext<RibbonGalleryHost | undefined>(RIBBON_GALLERY_HOST_KEY);
}

/** A context map carrying `host`, for `mount(Component, { context })` in tests. */
export function ribbonGalleryHostContext(host: RibbonGalleryHost): Map<symbol, unknown> {
	return new Map<symbol, unknown>([[RIBBON_GALLERY_HOST_KEY, host]]);
}
