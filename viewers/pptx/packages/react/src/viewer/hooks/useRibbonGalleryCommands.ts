import { tableLayoutCells } from 'ooxml-ui/pptx';
import type { RibbonGalleryApplyResult, RibbonGalleryContext } from 'ooxml-ui/pptx';
/**
 * useRibbonGalleryCommands: builds the {@link RibbonGalleryCommands} value
 * the viewer provides to every ribbon gallery.
 *
 * All decisions (what a gallery offers, what a pick writes) are shared; this
 * hook only gathers the context from viewer state and maps the three apply
 * result kinds onto the paths the inspector already uses: element patches go
 * through `updateElementById` (which records history via `markDirty`), theme
 * scheme picks through the theme editor's handlers.
 */
import type {
	PptxElement,
	PptxHandler,
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
	} = input;
	// Read at render time: the handler is created by the load pipeline, which
	// also replaces `theme`, so a new handler always comes with a re-render.
	const handler = handlerRef.current;
	const context = useMemo<RibbonGalleryContext>(
		() => ({
			element: selectedElement ?? null,
			themeColorMap,
			theme,
			resolveStyleMatrix: handler
				? (styleXml) => handler.resolveStyleMatrixReferences(styleXml)
				: undefined,
			selectedTableCells: tableLayoutCells(tableEditorState),
		}),
		[selectedElement, themeColorMap, theme, handler, tableEditorState],
	);
	const dispatch = useCallback(
		(result: RibbonGalleryApplyResult) => {
			if (!editable) {
				return;
			}
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
		[editable, updateElementById, updateThemeColorScheme, updateThemeFontScheme, applyThemeData],
	);
	return useMemo(() => ({ context, editable, dispatch }), [context, editable, dispatch]);
}
