import { commitsThroughDeckStyle, galleryDeckView, tableLayoutCells } from 'ooxml-ui/pptx';
import type { DeckStyleHost, RibbonGalleryApplyResult, RibbonGalleryContext } from 'ooxml-ui/pptx';
/**
 * useRibbonGalleryCommands: builds the {@link RibbonGalleryCommands} value
 * the viewer provides to every ribbon gallery.
 *
 * All decisions (what a gallery offers, what a pick writes) are shared; this
 * hook only gathers the context from viewer state and maps the apply result
 * kinds onto the viewer's paths: element patches go through `updateElementById`
 * (which records history via `markDirty`); theme, colour and font scheme picks,
 * like the effect and background picks, through the Design galleries' shared
 * commit (`deckStyle`), and through the theme handlers only without one.
 */
import type {
	PptxElement,
	PptxHandler,
	PptxSlide,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
} from 'pptx-viewer-core';
import { useCallback, useMemo } from 'react';
import type { RefObject } from 'react';

import type { RibbonGalleryCommands } from '../components/ribbon-gallery-context';
import type { TableCellEditorState } from '../types';

export interface UseRibbonGalleryCommandsInput {
	selectedElement: PptxElement | null | undefined;
	theme: PptxTheme | undefined;
	themeColorMap: Readonly<Record<string, string>> | undefined;
	handlerRef: RefObject<PptxHandler | null>;
	editable: boolean;
	/** The table cell cursor / range (Table Layout commands act on it). */
	tableEditorState?: TableCellEditorState | null;
	updateElementById: (elementId: string, updates: Partial<PptxElement>) => void;
	updateThemeColorScheme: (colorScheme: PptxThemeColorScheme) => Promise<void>;
	updateThemeFontScheme: (fontScheme: PptxThemeFontScheme) => Promise<void>;
	/** The slide being edited, for Background Styles' applied mark. */
	activeSlide?: PptxSlide | null | undefined;
	/** The Design galleries' hover preview and their effect and background commits. */
	deckStyle?: DeckStyleHost | undefined;
	/** Apply a whole theme (Design > Themes); without it a theme pick is ignored. */
	applyThemeData?: (
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme,
		themeName?: string,
	) => Promise<void>;
}

export function useRibbonGalleryCommands(
	input: UseRibbonGalleryCommandsInput,
): RibbonGalleryCommands {
	const {
		selectedElement,
		theme,
		themeColorMap,
		handlerRef,
		editable,
		tableEditorState,
		updateElementById,
		updateThemeColorScheme,
		updateThemeFontScheme,
		applyThemeData,
		activeSlide,
		deckStyle,
	} = input;
	// Read at render time: the handler is created by the load pipeline, which
	// also replaces `theme`, so a new handler always comes with a re-render.
	const handler = handlerRef.current;
	// While a hover preview shows, galleries are built from the original deck (stable tiles).
	const deck = galleryDeckView(deckStyle?.held(), { theme, themeColorMap, activeSlide });
	const context = useMemo<RibbonGalleryContext>(
		() => ({
			element: selectedElement ?? null,
			themeColorMap: deck.themeColorMap,
			theme: deck.theme,
			resolveStyleMatrix: handler
				? (styleXml, effectSchemeId) =>
						handler.resolveStyleMatrixReferences(styleXml, effectSchemeId)
				: undefined,
			selectedTableCells: tableLayoutCells(tableEditorState),
			activeSlide: deck.activeSlide ?? null,
			...(handler?.getThemeVariants().length ? { themeVariants: handler.getThemeVariants() } : {}),
		}),
		[selectedElement, deck.themeColorMap, deck.theme, handler, tableEditorState, deck.activeSlide],
	);
	const dispatch = useCallback(
		(result: RibbonGalleryApplyResult) => {
			if (!editable) {
				return;
			}
			if (deckStyle && commitsThroughDeckStyle(result)) {
				// Settles the hover preview itself, so the commit starts from the real slides. A theme
				// pick goes this way too: the pick arrives while its preview shows, and the theme
				// handlers would re-resolve from the theme of the last render, the PREVIEWED one, and
				// change nothing on the slides.
				void deckStyle.commit(result);
				return;
			}
			deckStyle?.endPreview();
			switch (result.kind) {
				case 'element':
					updateElementById(result.elementId, result.patch);
					return;
				case 'theme':
					void applyThemeData?.(result.colorScheme, result.fontScheme, result.name);
					return;
				case 'themeColorScheme':
					void updateThemeColorScheme(result.colorScheme);
					return;
				case 'themeFontScheme':
					void updateThemeFontScheme(result.fontScheme);
			}
		},
		[
			editable,
			updateElementById,
			updateThemeColorScheme,
			updateThemeFontScheme,
			applyThemeData,
			deckStyle,
		],
	);
	const preview = useCallback(
		(result: RibbonGalleryApplyResult | null) => {
			if (editable) {
				deckStyle?.preview(result);
			}
		},
		[editable, deckStyle],
	);
	const endPreview = useCallback(() => deckStyle?.endPreview(), [deckStyle]);
	return useMemo(
		() => ({ context, editable, dispatch, preview, endPreview }),
		[context, editable, dispatch, preview, endPreview],
	);
}
