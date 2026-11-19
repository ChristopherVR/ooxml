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
 * not torn down and rebuilt on every keystroke. Reads also see the port's own
 * writes before they render ({@link createPendingDeckWrites}), which is what lets
 * several proposals apply in one tick.
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

import { createPendingDeckWrites } from './pending-deck-writes';

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
	/** Changes on every presentation load (and room join) and only then: the bridge's `getDeckLoad`. */
	deckLoad: number;
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
	/** Record `next` as one undo step taken from `before`, ahead of its render. */
	recordSlidesStep: (next: PptxSlide[], label: string, before: PptxSlide[]) => void;
	markDirty: () => void;
	/** Route partial theme updates, made over `previous`, through the theme handlers. */
	applyThemeUpdates: (updates: Partial<PptxTheme>, previous: PptxTheme | undefined) => void;
	notify?: (message: string, level?: PptxAiNotifyLevel) => void;
}

export function useAiBridge(input: UseAiBridgeInput): PptxAiBridge {
	const ref = useRef(input);
	ref.current = input;

	return useMemo<PptxAiBridge>(() => {
		const live = (): UseAiBridgeInput => ref.current;
		const deckLoad = (): number => live().deckLoad;
		const deck = createPendingDeckWrites(live, deckLoad);
		return createPptxAiBridge({
			getSlides: () => deck.read('slides'),
			getActiveSlideIndex: () => live().activeSlideIndex,
			getCanvasSize: () => deck.read('canvasSize'),
			getTheme: () => deck.read('theme'),
			getHandler: () => live().handlerRef.current,
			getDeckLoad: deckLoad,
			getFileName: () => live().fileName,
			getSections: () => deck.read('sections'),
			getCustomShows: () => deck.read('customShows'),
			setCustomShows(shows) {
				deck.write('customShows', shows);
				live().setCustomShows(shows);
			},
			getPresentationProperties: () => deck.read('presentationProperties'),
			getCustomProperties: () => deck.read('customProperties'),
			getCoreProperties: () => deck.read('coreProperties'),
			getAppProperties: () => deck.read('appProperties'),
			getViewProperties: () => deck.read('viewProperties'),
			getTableStyleMap: () => deck.read('tableStyleMap'),
			getTableStylesDefaultId: () => deck.read('tableStylesDefaultId'),
			getTagCollections: () => deck.read('tagCollections'),
			goToSlide: (index) => live().setActiveSlideIndex(index),
			selectElements(slideIndex, elementIds) {
				const current = live();
				if (slideIndex !== current.activeSlideIndex) {
					current.setActiveSlideIndex(slideIndex);
				}
				current.applySelection(elementIds[0] ?? null, elementIds);
			},
			updateSlides(update, label) {
				const current = live();
				const before = deck.read('slides');
				const next = update(before);
				deck.write('slides', next);
				current.recordSlidesStep(next, label, before);
				// Another update queued in this tick (a theme recolour) is kept by
				// replaying this edit on top of it.
				current.setSlides((queued) => (queued === before ? next : update(queued)));
				current.markDirty();
			},
			applyTheme(updates) {
				const previous = deck.read('theme');
				deck.write('theme', { ...previous, ...updates });
				live().applyThemeUpdates(updates, previous);
			},
			setCanvasSize(size) {
				deck.write('canvasSize', size);
				live().setCanvasSize(size);
			},
			setSections(sections) {
				deck.write('sections', sections);
				live().setSections(sections);
			},
			setPresentationProperties(props) {
				deck.write('presentationProperties', props);
				live().setPresentationProperties(props);
			},
			setDocumentProperties(core, app, custom) {
				deck.write('coreProperties', core);
				deck.write('appProperties', app);
				deck.write('customProperties', custom);
				const current = live();
				current.setCoreProperties(core);
				current.setAppProperties(app);
				current.setCustomProperties(custom);
			},
			setViewProperties(props) {
				deck.write('viewProperties', props);
				live().setViewProperties(props);
			},
			setTableStyleMap(map) {
				deck.write('tableStyleMap', map);
				live().setTableStyleMap(map);
			},
			setTableStylesDefaultId(id) {
				deck.write('tableStylesDefaultId', id);
				live().setTableStylesDefaultId(id);
			},
			setTagCollections(tags) {
				deck.write('tagCollections', tags);
				live().setTagCollections(tags);
			},
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
