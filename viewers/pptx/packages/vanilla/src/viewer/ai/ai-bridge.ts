/**
 * {@link PptxAiBridge} implementation over the vanilla viewer controller.
 *
 * Reads come straight off the reactive store; navigation + selection reuse the
 * public viewer controls; and every write is funnelled through the editor's
 * history layer (`commitSlides` / `applyElementPatch`) so an AI edit is a single
 * undoable step, exactly like a manual one. This module has NO dependency on the
 * optional `ai` SDK, so the viewer can construct the bridge eagerly and only pay
 * for the SDK when the chat panel is actually opened.
 */

import type {
	PptxAiBridge,
	PptxAiBridgePort,
	PptxAiFocusedTarget,
	PptxAiNotifyLevel,
} from 'ooxml-ui/pptx/ai/host';
import { createPptxAiBridge } from 'ooxml-ui/pptx/ai/host';
import type { PptxCoreProperties, PptxHandler, PptxTheme } from 'pptx-viewer-core';

import type { EditorController } from '../editor';
import type { Store, ViewerState } from '../state';

export interface VanillaAiBridgeDeps {
	store: Store<ViewerState>;
	editor: EditorController;
	/** Navigate the viewer to a slide (viewer `goToSlide`). */
	goToSlide(index: number): void;
	/** Ensure editing is on so a write choke point is not silently dropped. */
	ensureEditable(): void;
	/** The loaded core handler, or null when nothing is loaded. */
	getHandler(): PptxHandler | null;
	/**
	 * Apply partial theme updates (colour/font scheme) to the deck and re-render.
	 * Theme scheme state lives outside the slides history snapshot, so this is a
	 * best-effort apply rather than a strictly undoable step.
	 */
	applyThemeUpdates(updates: Partial<PptxTheme>): void;
	/**
	 * The assistant's current focus (picks / pin / live selection), owned by the
	 * AI panel controller. When provided, it backs {@link PptxAiBridge.getFocusedTargets}
	 * so the context builder scopes the model to what the user is pointing at.
	 */
	getFocusedTargets?(): PptxAiFocusedTarget[];
	/** Called with the core properties when an AI edit changed them (the host's `onPropertiesChange`). */
	onCorePropertiesChanged?(properties: PptxCoreProperties): void;
	/** Optional host notification sink (status line / toast / console). */
	notify?(message: string, level?: PptxAiNotifyLevel): void;
}

/**
 * Build the AI bridge that exposes the live vanilla viewer to the AI core: the
 * shared `createPptxAiBridge` over a port that reads the store and writes
 * through the editor's undoable operations.
 */
export function createVanillaAiBridge(deps: VanillaAiBridgeDeps): PptxAiBridge {
	const { store, editor } = deps;
	const state = (): ViewerState => store.get();
	const port: PptxAiBridgePort = {
		getSlides: () => state().slides,
		getActiveSlideIndex: () => state().currentSlide,
		getCanvasSize: () => state().canvasSize,
		getTheme: () => {
			const { colorScheme } = state();
			return colorScheme ? { colorScheme } : undefined;
		},
		getHandler: () => deps.getHandler(),
		getSections: () => state().sections,
		getCustomShows: () => state().customShows,
		setCustomShows: (shows) => editor.updateCustomShows(shows),
		getPresentationProperties: () => state().presentationProperties,
		getCustomProperties: () => state().customProperties,
		getCoreProperties: () => state().coreProperties,
		getAppProperties: () => state().appProperties,
		getViewProperties: () => state().viewProperties,
		getTableStyleMap: () => state().tableStyleMap,
		getTableStylesDefaultId: () => state().tableStylesDefaultId,
		getTagCollections: () => state().tagCollections,
		goToSlide: (index) => deps.goToSlide(index),
		selectElements(slideIndex, elementIds) {
			deps.goToSlide(slideIndex);
			editor.selectElements(elementIds);
		},
		updateSlides(update) {
			const next = update(state().slides);
			deps.ensureEditable();
			editor.commitSlides(next, state().currentSlide);
		},
		applyTheme: (updates) => deps.applyThemeUpdates(updates),
		setCanvasSize: (size) => editor.getEditActions().updateCanvasSize(size),
		setSections: (sections) => editor.updateSections(sections),
		setPresentationProperties: (props) => editor.updatePresentationProperties(props),
		setDocumentProperties: (core, app, custom) =>
			editor.updateDocumentProperties(core, app, custom),
		// Not undo-tracked, mirroring the manual View > Grid/Guides/Snap toggle.
		setViewProperties: (viewProperties) => store.set({ viewProperties }),
		setTableStyleMap: (map) => editor.getEditActions().updateTableStyleMap(map),
		// No manual UI sets this yet, so it commits outside undo like viewProperties.
		setTableStylesDefaultId: (tableStylesDefaultId) => store.set({ tableStylesDefaultId }),
		setTagCollections: (tags) => editor.getEditActions().updateTagCollections(tags),
		onCorePropertiesChanged: (properties) => deps.onCorePropertiesChanged?.(properties),
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
