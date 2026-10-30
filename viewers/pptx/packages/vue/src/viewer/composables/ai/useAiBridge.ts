import type { CanvasSize } from 'ooxml-ui/pptx';
import {
	computeFocusTargets,
	createPptxAiBridge,
	resolveAiEffectiveTargets,
} from 'ooxml-ui/pptx/ai/host';
import type {
	PptxAiBridge,
	PptxAiBridgePort,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
} from 'ooxml-ui/pptx/ai/host';
/**
 * useAiBridge: builds a stable {@link PptxAiBridge} that lets the
 * framework-agnostic AI core (`pptx-viewer-shared/ai`) read the open deck,
 * navigate it, and route edits through the Vue editor-history layer so every AI
 * change is a single Ctrl+Z.
 *
 * The bridge itself is the shared `createPptxAiBridge`; this composable is only
 * its port over the Vue refs. Slide writes funnel through the same
 * `pushHistory()`-before-mutate path the manual editor uses, so an AI edit is
 * one undoable entry.
 *
 * The returned bridge object identity is stable (created once, closing over the
 * reactive refs), so the chat session built from it in {@link useAiChat} is not
 * torn down and rebuilt on every keystroke.
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
import type { Ref } from 'vue';

/** Live reactive inputs the bridge closes over. */
export interface UseAiBridgeInput {
	slides: Ref<PptxSlide[]>;
	activeSlideIndex: Ref<number>;
	canvasSize: Ref<CanvasSize>;
	theme: Ref<PptxTheme | undefined>;
	handler: Ref<PptxHandler | null>;
	// Presentation-level state, exposed so the AI's `getDeckData`/`applyDeckData`
	// seam can read and commit sections / canvas size / metadata / presentation
	// properties (the pptx-viewer-mcp "deck" tools). Slide + theme tools do not
	// need these; they route through applySlidesUpdate / applyTheme. Each is a
	// writable ref the bridge assigns `.value` on to fan a changed field out.
	sections: Ref<PptxSection[]>;
	customShows: Ref<PptxCustomShow[]>;
	presentationProperties: Ref<PptxPresentationProperties>;
	customProperties: Ref<PptxCustomProperty[]>;
	coreProperties: Ref<PptxCoreProperties | undefined>;
	/** The host's `properties-change` event: the core properties an AI edit set. */
	onCorePropertiesChanged?: (properties: PptxCoreProperties) => void;
	appProperties: Ref<PptxAppProperties | undefined>;
	/**
	 * `ppt/viewProps.xml` (grid/snap/guides toggles). Included so the deck
	 * tools' `getDeckData`/`applyDeckData` seam sees the same fields the main
	 * Save/Export path persists (`useLoadContent`'s `buildDeckSaveOptions` call).
	 */
	viewProperties: Ref<PptxViewProperties | undefined>;
	/** Parsed `ppt/tableStyles.xml` map, editable via the table style editor. */
	tableStyleMap: Ref<ParsedTableStyleMap | undefined>;
	/** `<a:tblStyleLst @def>` default style GUID. */
	tableStylesDefaultId: Ref<string | undefined>;
	/** `ppt/tags/*.xml` name/value metadata. */
	tagCollections: Ref<PptxTagCollection[]>;
	/** Display name of the open document, when known. */
	fileName: () => string | undefined;
	/** Snapshot the current slides onto the undo stack (call before mutating). */
	pushHistory: () => void;
	/** Flag the document dirty so autosave / status chrome updates. */
	markDirty: () => void;
	/** Navigate the viewer to a slide by zero-based index. */
	goTo: (index: number) => void;
	/** Replace the current element selection. */
	setSelection: (ids: string[]) => void;
	/** Route partial theme updates through the editor's theme handlers. */
	applyThemeUpdates: (updates: Partial<PptxTheme>) => void;
	/** All selected element ids on the active slide (drives live focus scope). */
	selectedElementIds?: () => string[];
	/** A pinned focus set from the chat (wins over the live selection). */
	pinnedFocus?: () => PptxAiFocusedTarget[] | null;
	/** Elements explicitly picked in pick mode (win over pin + selection). */
	pickedFocus?: () => PptxAiFocusedTarget[] | null;
	notify?: (message: string, level?: PptxAiNotifyLevel) => void;
}

export function useAiBridge(input: UseAiBridgeInput): PptxAiBridge {
	const port: PptxAiBridgePort = {
		getSlides: () => input.slides.value,
		getActiveSlideIndex: () => input.activeSlideIndex.value,
		getCanvasSize: () => input.canvasSize.value,
		getTheme: () => input.theme.value,
		getHandler: () => input.handler.value,
		getFileName: () => input.fileName(),
		getSections: () => input.sections.value,
		getCustomShows: () => input.customShows.value,
		setCustomShows: (shows) => {
			input.customShows.value = shows;
		},
		getPresentationProperties: () => input.presentationProperties.value,
		getCustomProperties: () => input.customProperties.value,
		getCoreProperties: () => input.coreProperties.value,
		getAppProperties: () => input.appProperties.value,
		getViewProperties: () => input.viewProperties.value,
		getTableStyleMap: () => input.tableStyleMap.value,
		getTableStylesDefaultId: () => input.tableStylesDefaultId.value,
		getTagCollections: () => input.tagCollections.value,
		goToSlide: (index) => input.goTo(index),
		selectElements(slideIndex, elementIds) {
			if (slideIndex !== input.activeSlideIndex.value) {
				input.goTo(slideIndex);
			}
			input.setSelection(elementIds);
		},
		// Push the current state, then install the updated slides: the explicit
		// push-before-mutate entry makes the AI edit a single Ctrl+Z.
		updateSlides(update) {
			input.pushHistory();
			input.slides.value = update(input.slides.value);
			input.markDirty();
		},
		applyTheme: (updates) => input.applyThemeUpdates(updates),
		setCanvasSize(size) {
			input.pushHistory();
			input.canvasSize.value = { width: size.width, height: size.height };
		},
		setSections: (sections) => {
			input.sections.value = sections;
		},
		setPresentationProperties: (props) => {
			input.presentationProperties.value = props;
		},
		setDocumentProperties(core, app, custom) {
			input.coreProperties.value = core;
			input.appProperties.value = app;
			input.customProperties.value = custom;
		},
		setViewProperties: (props) => {
			input.viewProperties.value = props;
		},
		setTableStyleMap: (map) => {
			input.tableStyleMap.value = map;
		},
		setTableStylesDefaultId: (id) => {
			input.tableStylesDefaultId.value = id;
		},
		setTagCollections: (tags) => {
			input.tagCollections.value = tags;
		},
		afterDeckUpdate: () => input.markDirty(),
		onCorePropertiesChanged: (properties) => input.onCorePropertiesChanged?.(properties),
		getFocusedTargets() {
			const selectedIds = input.selectedElementIds?.() ?? [];
			return resolveAiEffectiveTargets(
				{ pickTargets: input.pickedFocus?.() ?? [], pinnedFocus: input.pinnedFocus?.() ?? null },
				computeFocusTargets({
					activeSlideIndex: input.activeSlideIndex.value,
					selectedElementIds: selectedIds,
					selectedElementId: selectedIds[0] ?? null,
				}),
			);
		},
		notify: (message, level) => input.notify?.(message, level),
	};
	return createPptxAiBridge(port);
}
