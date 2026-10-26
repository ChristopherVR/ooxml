/**
 * {@link PptxAiBridge} implementation over the Svelte viewer's reactive editor.
 *
 * Reads come straight off getters the root component wires to its `EditorState`
 * / `PresentationLoader` runes; navigation + selection reuse the editor's public
 * ops; and every write is funnelled through {@link SvelteAiBridgeDeps.commitSlides}
 * so an AI edit becomes a single undoable history entry, exactly like a manual
 * one. This module has NO dependency on the optional `ai` SDK, so the root can
 * build the bridge eagerly and only pay for the SDK when the panel first opens.
 */

import type {
	PptxAiBridge,
	PptxAiBridgePort,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
} from 'ooxml-ui/pptx/ai/host';
import { createPptxAiBridge } from 'ooxml-ui/pptx/ai/host';
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

/** Live editor accessors the bridge closes over (all read from viewer runes). */
export interface SvelteAiBridgeDeps {
	/** The editable slides array (the single source of truth the AI mutates). */
	getSlides(): PptxSlide[];
	/** Zero-based index of the active slide. */
	getActiveSlideIndex(): number;
	/** Slide canvas size in CSS pixels. */
	getCanvasSize(): { width: number; height: number };
	/** The resolved presentation theme, when available. */
	getTheme(): PptxTheme | undefined;
	/** The loaded core handler, or null before a deck is open. */
	getHandler(): PptxHandler | null;
	/** Optional display file name, used as a friendly deck title. */
	getFileName(): string | undefined;
	/** Navigate the viewer to a slide by zero-based index. */
	goToSlide(index: number): void;
	/** Select elements on a slide (navigates first when off-slide). */
	selectElements(slideIndex: number, elementIds: string[]): void;
	/**
	 * Install a new slides array as ONE undoable history entry. The
	 * implementation is responsible for ensuring the editor is editable (so the
	 * commit is not silently dropped) before recording the step.
	 */
	commitSlides(next: PptxSlide[], label: string): void;
	/** Apply partial theme updates (colour/font scheme) and re-render. */
	applyTheme(updates: Partial<PptxTheme>): void;
	/** Optional host notification sink (status line / toast / console). */
	notify?(message: string, level?: PptxAiNotifyLevel): void;
	/**
	 * The assistant's focus (picks, pin, live selection) from the AI panel
	 * controller, so the model is scoped to what the user is pointing at.
	 */
	getFocusedTargets?(): PptxAiFocusedTarget[];

	// ── presentation-level (deck) state, for getDeckData / applyDeckData ───────
	// Svelte tracks the full deck: sections + document metadata live on the
	// editor (undoable), canvas size on the loader. These accessors/setters let
	// the AI's whole-deck seam read and commit the pptx-viewer-mcp "deck" tools.
	/** Deck sections (editor-tracked, undoable). */
	getSections(): PptxSection[];
	/** Deck-wide slide-show / print settings (editor-tracked, undoable). */
	getPresentationProperties(): PptxPresentationProperties;
	/** Document core properties (Title / Author / ...), when present. */
	getCoreProperties(): PptxCoreProperties | undefined;
	/** Application properties (Company / Application), when present. */
	getAppProperties(): PptxAppProperties | undefined;
	/** Custom document-property list (editor-tracked, undoable). */
	getCustomProperties(): PptxCustomProperty[];
	/**
	 * `ppt/viewProps.xml` (grid/snap/guides toggles), loader-tracked. Included
	 * so the deck tools' `getDeckData`/`applyDeckData` seam sees the same
	 * fields the main Save/Export path persists (`saveEditorDocument`).
	 */
	getViewProperties(): PptxViewProperties | undefined;
	/** Parsed `ppt/tableStyles.xml` map, loader-tracked. */
	getTableStyleMap(): ParsedTableStyleMap | undefined;
	/** `<a:tblStyleLst @def>` default style GUID, loader-tracked. */
	getTableStylesDefaultId(): string | undefined;
	/** `ppt/tags/*.xml` name/value metadata (editor-tracked, undoable). */
	getTagCollections(): PptxTagCollection[];
	/** Resize the slide canvas (loader value + editor history entry). */
	setCanvasSize(size: { width: number; height: number }): void;
	/** Replace the section list. */
	setSections(sections: PptxSection[]): void;
	/** Replace the deck-wide presentation properties. */
	setPresentationProperties(props: PptxPresentationProperties): void;
	/**
	 * Replace the document properties as one unit. Svelte's editor commits
	 * core / app / custom together, so the bridge fans all three through here.
	 */
	setDocumentProperties(
		core: PptxCoreProperties,
		app: PptxAppProperties,
		custom: PptxCustomProperty[],
	): void;
	/** Replace the view properties (grid/snap/guides toggles). Not undo-tracked, like the manual toggle. */
	setViewProperties(props: PptxViewProperties | undefined): void;
	/** Replace the table style map. Not undo-tracked, like the manual table style editor. */
	setTableStyleMap(map: ParsedTableStyleMap | undefined): void;
	/**
	 * Replace the default table style GUID. No manual UI sets this in the
	 * Svelte binding yet (a pre-existing gap, not introduced by this seam), so
	 * this is not undo-tracked, like `setTableStyleMap`.
	 */
	setTableStylesDefaultId(id: string | undefined): void;
	/** Replace the tag collections as one undoable edit. */
	setTagCollections(tags: PptxTagCollection[]): void;
}

/**
 * Build the AI bridge that exposes the live Svelte viewer to the AI core: the
 * shared `createPptxAiBridge` over these accessors.
 */
export function createSvelteAiBridge(deps: SvelteAiBridgeDeps): PptxAiBridge {
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
		updateSlides: (update, label) => deps.commitSlides(update(deps.getSlides()), label),
		applyTheme: (updates) => deps.applyTheme(updates),
		setCanvasSize: (size) => deps.setCanvasSize(size),
		setSections: (sections) => deps.setSections(sections),
		setPresentationProperties: (props) => deps.setPresentationProperties(props),
		setDocumentProperties: (core, app, custom) => deps.setDocumentProperties(core, app, custom),
		setViewProperties: (props) => deps.setViewProperties(props),
		setTableStyleMap: (map) => deps.setTableStyleMap(map),
		setTableStylesDefaultId: (id) => deps.setTableStylesDefaultId(id),
		setTagCollections: (tags) => deps.setTagCollections(tags),
		...(deps.getFocusedTargets
			? { getFocusedTargets: () => deps.getFocusedTargets?.() ?? [] }
			: {}),
		...(deps.notify
			? {
					notify: (message: string, level?: PptxAiNotifyLevel) => deps.notify?.(message, level),
				}
			: {}),
	};
	return createPptxAiBridge(port);
}
