import { queueThemeWrite } from 'ooxml-ui/pptx';
import type { HistoryTheme, ThemeEditorEdit } from 'ooxml-ui/pptx';
import {
	buildThemeColorMap,
	reResolveElementColors,
	reResolveElementFonts,
	reResolveSlideColors,
	reResolveSlideFonts,
} from 'pptx-viewer-core';
import type {
	PptxElement,
	PptxSlide,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeOption,
	PptxHandler,
} from 'pptx-viewer-core';
/**
 * useThemeHandlers: Theme application, colour-scheme / font-scheme / name
 * updates, presentation-wide theme apply, and template background handling.
 */
import type { RefObject } from 'react';

import { applyThemePartInPlace } from './theme-by-path-apply';
import { applyThemeEditorEdit } from './theme-editor-apply';
import { restoreHistoryTheme } from './theme-history-restore';
import type { EditorHistoryResult } from './useEditorHistory';

export interface UseThemeHandlersInput {
	handlerRef: RefObject<PptxHandler | null>;
	setTheme: React.Dispatch<React.SetStateAction<Record<string, unknown> | null>>;
	setSlideMasters: React.Dispatch<React.SetStateAction<Array<Record<string, unknown>>>>;
	slideMasters: Array<Record<string, unknown>>;
	/** The deck's theme parts as loaded; an undo back to the loaded theme restores its name. */
	themeOptions?: PptxThemeOption[];
	history: EditorHistoryResult;
	/** The live slides, which a theme applied by path re-themes in place. */
	slides: PptxSlide[];
	/** Current live slides, re-coloured in place on a scheme-colour change. */
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	/**
	 * Master/layout elements rendered as a separate per-slide overlay layer
	 * (`slide.elements` does not include them), re-coloured in place alongside
	 * `slides` so a scheme-colour change doesn't leave inherited background
	 * shapes painted with the old scheme until a full reload.
	 */
	templateElementsBySlideId: Record<string, PptxElement[]>;
	setTemplateElementsBySlideId: React.Dispatch<React.SetStateAction<Record<string, PptxElement[]>>>;
	/** Current theme, used to derive the previous colour map for re-resolution. */
	theme: PptxTheme | undefined;
	/**
	 * Force the history stack to capture the next slides change so an in-place
	 * theme re-colour lands as a single undoable entry (mirrors the AI path).
	 */
	bumpHistory: () => void;
}

export interface ThemeHandlersResult {
	handleApplyThemeEdit: (edit: ThemeEditorEdit) => Promise<void>;
	handleApplyTheme: (themePath: string, applyToAllMasters: boolean) => Promise<void>;
	handleUpdateThemeColorScheme: (colorScheme: PptxThemeColorScheme) => Promise<void>;
	handleUpdateThemeFontScheme: (fontScheme: PptxThemeFontScheme) => Promise<void>;
	handleUpdateThemeName: (name: string) => Promise<void>;
	handleApplyThemeToPresentation: () => Promise<void>;
	handleApplyThemeData: (
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme,
		themeName?: string,
	) => Promise<void>;
	/** Undo or redo crossed a theme change: put that theme back (see `useEditorHistory`). */
	handleRestoreHistoryTheme: (theme: HistoryTheme) => Promise<void>;
	handleSetTemplateBackground: (path: string, backgroundColor: string | undefined) => void;
	handleGetTemplateBackgroundColor: (path: string) => string | undefined;
	/** Return theme parts discovered in the loaded PPTX archive. */
	handleGetAvailableThemes: () => Promise<PptxThemeOption[]>;
	/**
	 * Switch every slide master to an existing theme part by its archive path, in place, as one
	 * undo step.
	 */
	handleSwitchTheme: (themePath: string) => Promise<void>;
}

export function useThemeHandlers(input: UseThemeHandlersInput): ThemeHandlersResult {
	const {
		handlerRef,
		setTheme,
		setSlideMasters,
		slideMasters,
		history,
		setSlides,
		templateElementsBySlideId,
		setTemplateElementsBySlideId,
		theme,
		bumpHistory,
	} = input;

	// Presentation properties > Theme: in place, as one undo step (never by reloading the deck).
	const handleApplyTheme = (themePath: string, applyToAllMasters: boolean) =>
		applyThemePartInPlace(input, themePath, applyToAllMasters);

	const handleUpdateThemeColorScheme = async (colorScheme: PptxThemeColorScheme) => {
		const handler = handlerRef.current;
		if (!handler) {
			return;
		}
		// Update the theme in the underlying archive + colour map so a later save
		// (and any future parse) carries the new scheme.
		await queueThemeWrite(handler, () => handler.updateThemeColorScheme(colorScheme));
		// Re-colour the live slides IN PLACE rather than serialising the whole
		// deck and feeding it back through the load pipeline. The old round-trip
		// (`serialize -> setContent -> full re-parse`) is a heavy main-thread
		// operation that also reset history/selection; driven by the AI mid-stream
		// or by the theme editor's colour picker (which fires continuously while
		// dragging) it stampeded into a multi-second renderer freeze. A direct
		// scheme-colour remap is cheap, keeps selection + history intact, and lands
		// as a single undoable entry.
		const previousMap = theme?.colorScheme ? buildThemeColorMap(theme.colorScheme) : {};
		setSlides((prev) => reResolveSlideColors(prev, previousMap, colorScheme));
		if (Object.keys(templateElementsBySlideId).length > 0) {
			setTemplateElementsBySlideId((prev) => {
				const next: Record<string, PptxElement[]> = {};
				for (const [slideId, elements] of Object.entries(prev)) {
					next[slideId] = reResolveElementColors(elements, previousMap, colorScheme);
				}
				return next;
			});
		}
		setTheme((prev) => (prev ? { ...prev, colorScheme } : { colorScheme }));
		bumpHistory();
		history.markDirty();
	};

	const handleUpdateThemeFontScheme = async (fontScheme: PptxThemeFontScheme) => {
		const handler = handlerRef.current;
		if (!handler) {
			return;
		}
		await queueThemeWrite(handler, () => handler.updateThemeFontScheme(fontScheme));
		// Re-resolve the live slides in place, as the colour scheme does: a full
		// reload would jump back to the first slide and drop the undo history.
		const previousFonts = theme?.fontScheme;
		bumpHistory();
		setSlides((prev) => reResolveSlideFonts(prev, previousFonts, fontScheme));
		setTemplateElementsBySlideId((prev) =>
			Object.fromEntries(
				Object.entries(prev).map(([slideId, elements]) => [
					slideId,
					reResolveElementFonts(elements, previousFonts, fontScheme),
				]),
			),
		);
		setTheme((prev) => (prev ? { ...prev, fontScheme } : { fontScheme }));
		history.markDirty();
	};

	const handleUpdateThemeName = async (name: string) => {
		const handler = handlerRef.current;
		if (!handler) {
			return;
		}
		await queueThemeWrite(handler, () => handler.updateThemeName(name));
		setTheme((prev) => (prev ? { ...prev, name } : { name }));
		history.markDirty();
	};

	// Every master on the first master's theme part.
	const handleApplyThemeToPresentation = async () => {
		const primaryThemePath = (slideMasters[0] as { themePath?: string } | undefined)?.themePath;
		if (primaryThemePath) {
			await applyThemePartInPlace(input, primaryThemePath, true);
		}
	};

	// Design > Themes and the theme gallery: re-themed in place (one undo step,
	// the active slide and selection kept), never by reloading the deck.
	const handleApplyThemeData = (
		colorScheme: PptxThemeColorScheme,
		fontScheme: PptxThemeFontScheme,
		themeName?: string,
	) =>
		applyThemeEditorEdit(input, { colorScheme, fontScheme, name: themeName ?? theme?.name ?? '' });

	const handleSetTemplateBackground = (path: string, backgroundColor: string | undefined) => {
		const handler = handlerRef.current;
		if (!handler) {
			return;
		}
		handler.setTemplateBackground(path, backgroundColor);
		setSlideMasters((prev) =>
			prev.map((m) => ((m as { path?: string }).path === path ? { ...m, backgroundColor } : m)),
		);
		history.markDirty();
	};

	const handleGetTemplateBackgroundColor = (path: string): string | undefined => {
		const handler = handlerRef.current;
		if (!handler) {
			return undefined;
		}
		return handler.getTemplateBackgroundColor(path);
	};

	const handleGetAvailableThemes = async (): Promise<PptxThemeOption[]> => {
		const handler = handlerRef.current;
		if (!handler) {
			return [];
		}
		return handler.getAvailableThemes();
	};

	const handleSwitchTheme = (themePath: string): Promise<void> =>
		applyThemePartInPlace(input, themePath, true);

	const handleApplyThemeEdit = (edit: ThemeEditorEdit) => applyThemeEditorEdit(input, edit);
	const handleRestoreHistoryTheme = (target: HistoryTheme) => restoreHistoryTheme(input, target);

	return {
		handleApplyThemeEdit,
		handleApplyTheme,
		handleUpdateThemeColorScheme,
		handleUpdateThemeFontScheme,
		handleUpdateThemeName,
		handleApplyThemeToPresentation,
		handleApplyThemeData,
		handleRestoreHistoryTheme,
		handleSetTemplateBackground,
		handleGetTemplateBackgroundColor,
		handleGetAvailableThemes,
		handleSwitchTheme,
	};
}
