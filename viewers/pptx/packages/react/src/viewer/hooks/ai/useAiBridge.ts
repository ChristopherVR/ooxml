import type { PptxAiBridge, PptxAiFocusedTarget, PptxAiNotifyLevel } from 'ooxml-ui/pptx/ai/host';
import {
	computeFocusTargets,
	createPptxAiBridge,
	resolveAiEffectiveTargets,
} from 'ooxml-ui/pptx/ai/host';
/**
 * useAiBridge: builds a stable {@link PptxAiBridge} that lets the
 * framework-agnostic AI core (`pptx-viewer-shared/ai`) read the open deck,
 * navigate it, and route edits through the React editor-history layer so every
 * AI change is a single Ctrl+Z.
 *
 * The bridge is the shared `createPptxAiBridge`; this hook is only its port over
 * React state. The bridge identity is kept stable across renders (the port
 * reads live values from a ref) so the chat session built from it in {@link useAiChat} is
 * not torn down and rebuilt on every keystroke.
 */
import type {
	ParsedTableStyleMap,
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxHandler,
	PptxPresentationProperties,
	PptxCustomShow,
	PptxSection,
	PptxSlide,
	PptxTagCollection,
	PptxTheme,
	PptxViewProperties,
} from 'pptx-viewer-core';
import type { RefObject } from 'react';
import { useMemo, useRef } from 'react';

/** Live inputs the bridge closes over. Updated on every render via a ref. */
export interface UseAiBridgeInput {
	slides: PptxSlide[];
	activeSlideIndex: number;
	canvasSize: { width: number; height: number };
	theme: PptxTheme | undefined;
	fileName?: string;
	/** Primary selected element id on the active slide (null when none). */
	selectedElementId: string | null;
	/** All selected element ids on the active slide (multi-select, tables incl). */
	selectedElementIds: string[];
	/**
	 * When the user pins a focus in the chat, these override the live selection
	 * for {@link PptxAiBridge.getFocusedTargets}; null means "follow selection".
	 */
	pinnedFocus: PptxAiFocusedTarget[] | null;
	/**
	 * Elements the user explicitly picked (pick mode). When non-empty these are
	 * the assistant's focus, winning over a pin and the live selection.
	 */
	pickedFocus?: PptxAiFocusedTarget[] | null;
	handlerRef: RefObject<PptxHandler | null>;
	// Presentation-level state, exposed so the AI's `getDeckData`/`applyDeckData`
	// seam can read and commit metadata / sections / canvas size / presentation
	// properties (the pptx-viewer-mcp "deck" tools). Slide + theme tools do not
	// need these; they route through applySlidesUpdate / applyTheme.
	sections: PptxSection[];
	presentationProperties: PptxPresentationProperties;
	customProperties: PptxCustomProperty[];
	coreProperties: PptxCoreProperties | undefined;
	appProperties: PptxAppProperties | undefined;
	/**
	 * `ppt/viewProps.xml` (grid/snap/guides toggles). Included so the deck
	 * tools' `getDeckData`/`applyDeckData` seam sees the same fields the main
	 * Save/Export path persists (`useSerialize`'s `buildDeckSaveOptions` call).
	 */
	viewProperties: PptxViewProperties | undefined;
	/** Parsed `ppt/tableStyles.xml` map, editable via the table style editor. */
	tableStyleMap: ParsedTableStyleMap | undefined;
	/** `<a:tblStyleLst @def>` default style GUID. */
	tableStylesDefaultId: string | undefined;
	/** `ppt/tags/*.xml` name/value metadata. */
	tagCollections: PptxTagCollection[];
	setCanvasSize: (size: { width: number; height: number }) => void;
	setSections: (sections: PptxSection[]) => void;
	customShows: PptxCustomShow[];
	setCustomShows: (shows: PptxCustomShow[]) => void;
	setPresentationProperties: (props: PptxPresentationProperties) => void;
	setCustomProperties: (props: PptxCustomProperty[]) => void;
	setCoreProperties: (props: PptxCoreProperties | undefined) => void;
	setAppProperties: (props: PptxAppProperties | undefined) => void;
	setViewProperties: (props: PptxViewProperties | undefined) => void;
	setTableStyleMap: (map: ParsedTableStyleMap | undefined) => void;
	setTableStylesDefaultId: (id: string | undefined) => void;
	setTagCollections: (tags: PptxTagCollection[]) => void;
	setSlides: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
	setActiveSlideIndex: (index: number) => void;
	applySelection: (primaryId: string | null, ids?: string[]) => void;
	/** Force the history stack to capture the next slides change (content-only edits). */
	bumpHistory: () => void;
	markDirty: () => void;
	/** Route partial theme updates through the editor's theme handlers. */
	applyThemeUpdates: (updates: Partial<PptxTheme>) => void;
	notify?: (message: string, level?: PptxAiNotifyLevel) => void;
}

export function useAiBridge(input: UseAiBridgeInput): PptxAiBridge {
	const ref = useRef(input);
	ref.current = input;

	return useMemo<PptxAiBridge>(() => {
		const live = (): UseAiBridgeInput => ref.current;
		return createPptxAiBridge({
			getSlides: () => live().slides,
			getActiveSlideIndex: () => live().activeSlideIndex,
			getCanvasSize: () => live().canvasSize,
			getTheme: () => live().theme,
			getHandler: () => live().handlerRef.current,
			getFileName: () => live().fileName,
			getSections: () => live().sections,
			getCustomShows: () => live().customShows,
			setCustomShows: (shows) => live().setCustomShows(shows),
			getPresentationProperties: () => live().presentationProperties,
			getCustomProperties: () => live().customProperties,
			getCoreProperties: () => live().coreProperties,
			getAppProperties: () => live().appProperties,
			getViewProperties: () => live().viewProperties,
			getTableStyleMap: () => live().tableStyleMap,
			getTableStylesDefaultId: () => live().tableStylesDefaultId,
			getTagCollections: () => live().tagCollections,
			goToSlide: (index) => live().setActiveSlideIndex(index),
			selectElements(slideIndex, elementIds) {
				const current = live();
				if (slideIndex !== current.activeSlideIndex) {
					current.setActiveSlideIndex(slideIndex);
				}
				current.applySelection(elementIds[0] ?? null, elementIds);
			},
			updateSlides(update) {
				const current = live();
				// A functional update, so two AI edits in one tick compose.
				current.setSlides((prev) => update(prev));
				// Content-only edits (text/style) do not change slide/element counts,
				// so nudge the history nonce to force the snapshot capture that makes
				// the AI edit a single undoable entry.
				current.bumpHistory();
				current.markDirty();
			},
			applyTheme: (updates) => live().applyThemeUpdates(updates),
			setCanvasSize: (size) => live().setCanvasSize(size),
			setSections: (sections) => live().setSections(sections),
			setPresentationProperties: (props) => live().setPresentationProperties(props),
			setDocumentProperties(core, app, custom) {
				const current = live();
				current.setCoreProperties(core);
				current.setAppProperties(app);
				current.setCustomProperties(custom);
			},
			setViewProperties: (props) => live().setViewProperties(props),
			setTableStyleMap: (map) => live().setTableStyleMap(map),
			setTableStylesDefaultId: (id) => live().setTableStylesDefaultId(id),
			setTagCollections: (tags) => live().setTagCollections(tags),
			afterDeckUpdate: () => live().markDirty(),
			getFocusedTargets() {
				const current = live();
				// Explicit picks win, then a pinned focus, then the live selection.
				return resolveAiEffectiveTargets(
					{ pickTargets: current.pickedFocus ?? [], pinnedFocus: current.pinnedFocus },
					computeFocusTargets({
						activeSlideIndex: current.activeSlideIndex,
						selectedElementIds: current.selectedElementIds,
						selectedElementId: current.selectedElementId,
					}),
				);
			},
			notify: (message, level) => live().notify?.(message, level),
		});
	}, []);
}
