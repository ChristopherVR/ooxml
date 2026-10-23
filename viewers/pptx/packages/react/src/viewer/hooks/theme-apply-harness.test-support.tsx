/**
 * A viewer reduced to what a theme pick touches, for `theme-apply-in-place.repro.test.tsx`.
 *
 * Wires the REAL `useLoadContent`, `useEditorHistory`, `useSerialize`, `useThemeHandlers`,
 * `useDeckStyleHost` and `useRibbonGalleryCommands` the way `PowerPointViewer` does, and exposes what a reload would
 * reset (the parse count, the selection, the slide on screen, the dirty flag, the undo stack) plus
 * the theme and the saved bytes. Test-only: nothing in the package imports it.
 */
import type {
	PptxElement,
	PptxHandler,
	PptxSlide,
	PptxSlideMaster,
	PptxTheme,
	PptxThemeOption,
} from 'pptx-viewer-core';
import React, { act, useCallback, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';

import type { RibbonGalleryCommands } from '../components/ribbon-gallery-context';
import type { CanvasSize, EditorHistoryTheme } from '../types';
import { useDeckStyleHost } from './useDeckStyleHost';
import { useEditorHistory } from './useEditorHistory';
import type { EditorHistoryResult } from './useEditorHistory';
import { useLoadContent } from './useLoadContent';
import type { UseLoadContentInput } from './useLoadContent';
import { useRibbonGalleryCommands } from './useRibbonGalleryCommands';
import { useSerialize } from './useSerialize';
import { useThemeHandlers } from './useThemeHandlers';
import type { ThemeHandlersResult, UseThemeHandlersInput } from './useThemeHandlers';

/** The load setters this harness does not observe. */
const INERT_LOAD_SETTERS = Object.fromEntries(
	`setSlideSizeEmu setHeaderFooter setFirstSlideNumber setLayoutOptions setModernCommentAuthors
	setRecentColors setTableStyleMap setTableStylesDefaultId setTableStylesToDelete setCustomShows
	setActiveCustomShowId setSections setPresentationProperties setViewProperties setNotesMaster
	setHandoutMaster setNotesCanvasSize setCustomProperties setTagCollections setCoreProperties
	setAppProperties setEmbeddedFonts setHasMacros setHasDigitalSignatures
	setDigitalSignatureCount setGuides setReadOnlyRecommendation setModifyVerifier
	setCompatToasts setError setIsEncrypted`
		.split(/\s+/u)
		.map((name) => [name, () => {}]),
) as unknown as UseLoadContentInput;

const NO_GUIDES: Array<{ id: string; axis: 'h' | 'v'; position: number }> = [];
const NO_HEADER_FOOTER = {};
const NO_PRESENTATION_PROPERTIES = {};
const NO_LIST: never[] = [];

const noPointer = (): boolean => false;
const noop = (): void => {};

export interface Probe {
	initial: Uint8Array;
	content: ArrayBuffer | Uint8Array | null;
	parses: number;
	selectionClears: number;
	slides: PptxSlide[];
	activeSlideIndex: number;
	isDirty: boolean;
	loading: boolean;
	theme: PptxTheme | undefined;
	history: EditorHistoryResult;
	gallery: RibbonGalleryCommands;
	/** The theme handlers Browse Themes, the theme editor and the assistant commit through. */
	themeHandlers: ThemeHandlersResult;
	setActiveSlideIndex: (index: number) => void;
	/** The loaded handler (the theme parts live in it). */
	handler: () => PptxHandler | null;
	/** The deck as the viewer's save path (`useSerialize`) writes it now. */
	save: () => Promise<Uint8Array>;
}

let probe: Probe | null = null;

export function live(): Probe {
	if (!probe) {
		throw new Error('the harness has not rendered');
	}
	return probe;
}

function Harness({ initial }: { initial: Uint8Array }): React.ReactElement {
	const [content, setContent] = useState<ArrayBuffer | Uint8Array | null>(initial);
	const [slides, setSlides] = useState<PptxSlide[]>([]);
	const [templates, setTemplates] = useState<Record<string, PptxElement[]>>({});
	const [theme, setTheme] = useState<PptxTheme | undefined>(undefined);
	const [themeOptions, setThemeOptions] = useState<PptxThemeOption[]>([]);
	const [slideMasters, setSlideMasters] = useState<PptxSlideMaster[]>([]);
	const [activeSlideIndex, setActiveSlideIndex] = useState(0);
	const [canvasSize, setCanvasSize] = useState<CanvasSize>({ width: 1280, height: 720 });
	const [isDirty, setIsDirty] = useState(false);
	const [loading, setLoading] = useState(true);
	const [pointerCommitNonce, setPointerCommitNonce] = useState(0);
	const parses = useRef(0);
	const selectionClears = useRef(0);
	const restoreTheme = useRef<((target: EditorHistoryTheme) => void) | undefined>(undefined);
	const mediaDataUrls = useMemo(() => new Map<string, string>(), []);
	const markDocumentDirty = useCallback(() => setIsDirty(true), []);
	const bumpHistory = useCallback(() => setPointerCommitNonce((nonce) => nonce + 1), []);

	const history = useEditorHistory({
		slides,
		canvasSize,
		activeSlideIndex,
		templateElementsBySlideId: templates,
		selectedElementId: null,
		selectedElementIds: [],
		editTemplateMode: false,
		headerFooter: {},
		loading,
		error: null,
		hasActivePointerInteraction: noPointer,
		pointerCommitNonce,
		onDirty: markDocumentDirty,
		theme,
		onRestoreTheme: (target) => restoreTheme.current?.(target),
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId: setTemplates,
		setSelectedElementId: noop,
		setSelectedElementIds: noop,
		setEditTemplateMode: noop,
		setHeaderFooter: noop,
	});
	const { handlerRef } = useLoadContent({
		...INERT_LOAD_SETTERS,
		content,
		clearSelection: () => {
			selectionClears.current += 1;
		},
		history,
		setSlides,
		setTemplateElementsBySlideId: setTemplates,
		mediaDataUrls,
		setCanvasSize,
		setTheme,
		setThemeOptions,
		setSlideMasters,
		setActiveSlideIndex,
		setLoading,
		setIsDirty,
		onContentApplied: () => {
			parses.current += 1;
		},
	});
	// The viewer's own save path: Save, Save As, export, autosave and `getContent()` all go through it.
	const inlineEditingElementIdRef = useRef<string | null>(null);
	const inlineEditingTextRef = useRef('');
	const serializeSlides = useSerialize({
		slides,
		templateElementsBySlideId: templates,
		activeSlideIndex,
		canvasSize,
		guides: NO_GUIDES,
		headerFooter: NO_HEADER_FOOTER,
		presentationProperties: NO_PRESENTATION_PROPERTIES,
		customShows: NO_LIST,
		sections: NO_LIST,
		coreProperties: undefined,
		appProperties: undefined,
		customProperties: NO_LIST,
		tagCollections: NO_LIST,
		slideMasters,
		notesMaster: undefined,
		handoutMaster: undefined,
		handlerRef,
		inlineEditingElementIdRef,
		inlineEditingTextRef,
	});
	const themeHandlers = useThemeHandlers({
		handlerRef,
		serializeSlides,
		setContent,
		onContentChange: undefined,
		setTheme: setTheme as unknown as UseThemeHandlersInput['setTheme'],
		setSlideMasters: noop,
		slideMasters: slideMasters as unknown as UseThemeHandlersInput['slideMasters'],
		themeOptions,
		history,
		setSlides,
		templateElementsBySlideId: templates,
		setTemplateElementsBySlideId: setTemplates,
		theme,
		bumpHistory,
	});
	restoreTheme.current = (target) => void themeHandlers.handleRestoreHistoryTheme(target);
	const deckStyle = useDeckStyleHost({
		slides,
		templateElementsBySlideId: templates,
		theme,
		themeColorMap: undefined,
		handlerRef,
		selectedSlideIds: slides[activeSlideIndex] ? [slides[activeSlideIndex].id] : [],
		setSlides,
		setTemplateElementsBySlideId: setTemplates,
		setTheme,
		holdTracking: history.holdTracking,
		bumpHistory,
		markDirty: history.markDirty,
	});
	const gallery = useRibbonGalleryCommands({
		deckStyle,
		selectedElement: null,
		theme,
		themeColorMap: undefined,
		handlerRef,
		editable: true,
		updateElementById: noop,
		updateThemeColorScheme: themeHandlers.handleUpdateThemeColorScheme,
		updateThemeFontScheme: themeHandlers.handleUpdateThemeFontScheme,
		applyThemeData: themeHandlers.handleApplyThemeData,
	});

	probe = {
		initial,
		content,
		parses: parses.current,
		selectionClears: selectionClears.current,
		slides,
		activeSlideIndex,
		isDirty,
		loading,
		theme,
		history,
		gallery,
		themeHandlers,
		setActiveSlideIndex,
		handler: () => handlerRef.current,
		save: async () => {
			const bytes = await serializeSlides();
			if (!bytes) {
				throw new Error('no deck is loaded');
			}
			return bytes;
		},
	};
	return React.createElement('div');
}

export async function flush(ms = 0): Promise<void> {
	await act(async () => {
		await new Promise<void>((done) => {
			setTimeout(done, ms);
		});
	});
}

export async function flushUntil(isDone: () => boolean, timeoutMs = 8000): Promise<void> {
	const deadline = Date.now() + timeoutMs;
	while (Date.now() < deadline && !isDone()) {
		await flush(10);
	}
}

/** The face the slide-2 title ("Beta") is drawn in. */
export function betaTitleFont(): string | undefined {
	const title = live().slides[1]?.elements.find((element) => element.id.endsWith('shape-0'));
	return title && 'textSegments' in title ? title.textSegments?.[0]?.style?.fontFamily : undefined;
}

/** What a reload resets, read as one value so a failure shows every part at once. */
export function session(baseline: { parses: number; selectionClears: number }) {
	const current = live();
	return {
		reloaded: current.parses !== baseline.parses || current.content !== current.initial,
		selectionCleared: current.selectionClears !== baseline.selectionClears,
		slideOnScreen: current.activeSlideIndex + 1,
		dirty: current.isDirty,
		canUndo: current.history.canUndo,
	};
}

/** The flag React reads to know `act` is in use; set for the harness and put back after. */
const actGlobal = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };

/** Mount the harness over `deck`, run `body`, and unmount whatever happens. */
export async function withHarness(deck: Uint8Array, body: () => Promise<void>): Promise<void> {
	const hadActEnvironment = Object.hasOwn(actGlobal, 'IS_REACT_ACT_ENVIRONMENT');
	const previousActEnvironment = actGlobal.IS_REACT_ACT_ENVIRONMENT;
	actGlobal.IS_REACT_ACT_ENVIRONMENT = true;
	const root = createRoot(document.createElement('div'));
	try {
		await act(async () => {
			root.render(React.createElement(Harness, { initial: deck }));
		});
		await body();
	} finally {
		await act(async () => {
			root.unmount();
		});
		probe = null;
		if (hadActEnvironment) {
			actGlobal.IS_REACT_ACT_ENVIRONMENT = previousActEnvironment;
		} else {
			delete actGlobal.IS_REACT_ACT_ENVIRONMENT;
		}
	}
}
