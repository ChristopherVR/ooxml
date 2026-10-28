import type {
	PptxAiBridge,
	PptxAiBridgePort,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
} from 'ooxml-ui/pptx/ai';
import { createPptxAiBridge } from 'ooxml-ui/pptx/ai/host';
/**
 * createAngularAiBridge: builds a {@link PptxAiBridge} that lets the
 * framework-agnostic AI core (`pptx-viewer-shared/ai`) read the open deck,
 * navigate it, and route edits through the Angular editor-history layer so
 * every AI change is a single, undoable Ctrl+Z.
 *
 * The bridge is a plain object closing over a small set of accessor callbacks
 * (mirroring React's `useAiBridge` ref indirection): each read pulls the live
 * value on demand, and each of the three write choke points funnels through
 * {@link BridgeDeps.applySlides} -> `EditorStateService.applyReplacement`, which
 * records ONE history snapshot per commit. Keeping writes funnelled is what lets
 * the {@link ProposalStore} apply a staged batch atomically.
 */
import type {
	ParsedTableStyleMap,
	PptxAppProperties,
	PptxCoreProperties,
	PptxCustomProperty,
	PptxHandler,
	PptxPresentationProperties,
	PptxSection,
	PptxSlide,
	PptxTagCollection,
	PptxTheme,
	PptxViewProperties,
} from 'pptx-viewer-core';

/** Live accessors the bridge closes over, supplied by the viewer component. */
export interface BridgeDeps {
	/** The editable (template-free) slide deck. */
	getSlides(): readonly PptxSlide[];
	/** Zero-based index of the active slide. */
	getActiveSlideIndex(): number;
	/** Slide canvas size in CSS pixels. */
	getCanvasSize(): { width: number; height: number };
	/** The resolved presentation theme, when loaded. */
	getTheme(): PptxTheme | undefined;
	/** Host-provided document display name, when set. */
	getFileName(): string | undefined;
	/** The loaded core handler, when a deck is open. */
	getHandler(): PptxHandler | undefined;
	/** Navigate the viewer to a slide by zero-based index. */
	goToSlide(index: number): void;
	/** Select the given elements on a slide (empty clears selection). */
	selectElements(slideIndex: number, elementIds: readonly string[]): void;
	/**
	 * Commit a fully-computed next slides array as ONE undoable history entry.
	 * Implemented by the component as `editor.applyReplacement(next, label)`.
	 */
	applySlides(next: PptxSlide[], label: string): void;
	/** Apply partial theme updates through the editor's theme handlers. */
	applyTheme(updates: Partial<PptxTheme>): void;
	// ── presentation-level (deck) state, for the AI `getDeckData`/`applyDeckData`
	// seam that the pptx-viewer-mcp "deck" tools (canvas size, metadata, sections,
	// presentation properties) route through. Slide/theme tools do not need these.
	/** Ordered presentation sections (editor-tracked, saved with the deck). */
	getSections(): readonly PptxSection[];
	/** Presentation-level properties (`p:presentationPr`), loader-tracked. */
	getPresentationProperties(): PptxPresentationProperties;
	/** Custom document properties (`docProps/custom.xml`), loader-tracked. */
	getCustomProperties(): readonly PptxCustomProperty[];
	/** Core document properties (`docProps/core.xml`), loader-tracked. */
	getCoreProperties(): PptxCoreProperties | undefined;
	/** Extended application properties (`docProps/app.xml`), loader-tracked. */
	getAppProperties(): PptxAppProperties | undefined;
	/**
	 * `ppt/viewProps.xml` (grid/snap/guides toggles), loader-tracked. Included
	 * so the deck tools' `getDeckData`/`applyDeckData` seam sees the same
	 * fields the main Save/Export path persists (`loader.saveSlides`).
	 */
	getViewProperties(): PptxViewProperties | undefined;
	/** Parsed `ppt/tableStyles.xml` map, loader-tracked. */
	getTableStyleMap(): ParsedTableStyleMap | undefined;
	/** `<a:tblStyleLst @def>` default style GUID, loader-tracked. */
	getTableStylesDefaultId(): string | undefined;
	/** `ppt/tags/*.xml` name/value metadata, loader-tracked. */
	getTagCollections(): readonly PptxTagCollection[];
	/** Set the slide canvas size (px) and mark the deck dirty. */
	setCanvasSize(size: { width: number; height: number }): void;
	/** Replace the presentation sections and mark the deck dirty. */
	setSections(sections: readonly PptxSection[]): void;
	/** Replace the presentation-level properties and mark the deck dirty. */
	setPresentationProperties(props: PptxPresentationProperties): void;
	/** Replace the custom document properties and mark the deck dirty. */
	setCustomProperties(props: readonly PptxCustomProperty[]): void;
	/** Replace the core document properties and mark the deck dirty. */
	setCoreProperties(props: PptxCoreProperties | undefined): void;
	/** Replace the extended application properties and mark the deck dirty. */
	setAppProperties(props: PptxAppProperties | undefined): void;
	/** Replace the view properties and mark the deck dirty. */
	setViewProperties(props: PptxViewProperties | undefined): void;
	/** Replace the table style map and mark the deck dirty. */
	setTableStyleMap(map: ParsedTableStyleMap | undefined): void;
	/** Replace the default table style GUID and mark the deck dirty. */
	setTableStylesDefaultId(id: string | undefined): void;
	/** Replace the tag collections and mark the deck dirty. */
	setTagCollections(tags: readonly PptxTagCollection[]): void;
	/**
	 * The slides / elements the assistant should scope its work to (explicit AI
	 * picks, else a pinned focus, else the live canvas selection). When omitted,
	 * the assistant falls back to the whole active slide.
	 */
	getFocusedTargets?(): PptxAiFocusedTarget[];
	/** Optional transient host notification (toast / status line). */
	notify?(message: string, level?: PptxAiNotifyLevel): void;
}

/**
 * Build a stable {@link PptxAiBridge} over the Angular viewer state/editor: the
 * shared `createPptxAiBridge` over these accessors.
 */
export function createAngularAiBridge(deps: BridgeDeps): PptxAiBridge {
	const port: PptxAiBridgePort = {
		getSlides: () => deps.getSlides(),
		getActiveSlideIndex: () => deps.getActiveSlideIndex(),
		getCanvasSize: () => deps.getCanvasSize(),
		getTheme: () => deps.getTheme(),
		getHandler: () => deps.getHandler(),
		getFileName: () => deps.getFileName(),
		getSections: () => deps.getSections(),
		getPresentationProperties: () => deps.getPresentationProperties(),
		getCustomProperties: () => deps.getCustomProperties(),
		getCoreProperties: () => deps.getCoreProperties(),
		getAppProperties: () => deps.getAppProperties(),
		getViewProperties: () => deps.getViewProperties(),
		getTableStyleMap: () => deps.getTableStyleMap(),
		getTableStylesDefaultId: () => deps.getTableStylesDefaultId(),
		getTagCollections: () => deps.getTagCollections(),
		goToSlide: (index) => deps.goToSlide(index),
		selectElements: (slideIndex, elementIds) => deps.selectElements(slideIndex, elementIds),
		// One `applyReplacement` per commit, so an AI edit is one Ctrl+Z.
		updateSlides: (update, label) => deps.applySlides(update(deps.getSlides()), label),
		applyTheme: (updates) => deps.applyTheme(updates),
		setCanvasSize: (size) => deps.setCanvasSize(size),
		setSections: (sections) => deps.setSections(sections),
		setPresentationProperties: (props) => deps.setPresentationProperties(props),
		setDocumentProperties(core, app, custom) {
			deps.setCoreProperties(core);
			deps.setAppProperties(app);
			deps.setCustomProperties(custom);
		},
		setViewProperties: (props) => deps.setViewProperties(props),
		setTableStyleMap: (map) => deps.setTableStyleMap(map),
		setTableStylesDefaultId: (id) => deps.setTableStylesDefaultId(id),
		setTagCollections: (tags) => deps.setTagCollections(tags),
		// Without a panel focus the assistant scopes to the whole active slide.
		getFocusedTargets: () =>
			deps.getFocusedTargets?.() ?? [{ kind: 'slide', slideIndex: deps.getActiveSlideIndex() }],
		notify: (message, level) => deps.notify?.(message, level),
	};
	return createPptxAiBridge(port);
}
