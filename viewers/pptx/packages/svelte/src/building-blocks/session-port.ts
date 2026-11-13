import {
	BUILDING_BLOCK_OK,
	buildingBlockFailure,
	hyperlinkEditPatch,
	resetSlideLayoutPath,
	scopeLayoutOptionsToSlide,
	selectionEquation,
} from 'ooxml-ui/pptx';
import type {
	BuildingBlockAvailabilityInput,
	BuildingBlockFindReplacePort,
	BuildingBlockSessionPort,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxLayoutOption } from 'pptx-viewer-core';

import {
	buildChartInsertElement,
	buildEquationInsertElement,
	buildMediaInsertElement,
	buildSmartArtInsertElement,
} from '../viewer/editor';
import type { FindReplaceState } from '../viewer/editor/editor-find-replace.svelte';
import type { ViewerStateBag } from '../viewer/state/create-viewer-state-types';
import type { ViewerBuildingBlocksOptions } from './types';

/** What the Svelte port reads at call time. */
export interface SveltePortSource {
	state: ViewerStateBag;
	options(): ViewerBuildingBlocksOptions;
	/** The deck's layouts, loaded asynchronously by the blocks (reactive). */
	layouts(): readonly PptxLayoutOption[];
}

/** The first of the effective selection, as the insert handlers read it. */
export function selectedElementOf(state: ViewerStateBag): PptxElement | null {
	return state.editor.selectedElements[0] ?? null;
}

/** The layouts the active slide may switch to (its master's). */
export function scopedLayoutsOf(
	state: ViewerStateBag,
	layouts: readonly PptxLayoutOption[],
): PptxLayoutOption[] {
	return scopeLayoutOptionsToSlide(
		[...layouts],
		state.editor.slides[state.viewer.current]?.layoutPath,
	);
}

/** Whether the blocks may edit: the host's permission, collaboration and the deck's own locks. */
export function canEditOf(state: ViewerStateBag): boolean {
	return state.editable && !state.collab.readOnly && !state.readOnlyRec.locked;
}

/** The availability input of the Svelte viewer state, gathered from its runes. */
export function svelteAvailabilityInput(source: SveltePortSource): BuildingBlockAvailabilityInput {
	const { state } = source;
	const element = selectedElementOf(state);
	const slide = state.editor.slides[state.viewer.current];
	return {
		ready: !state.loader.loading && !state.loader.error,
		hasActiveSlide: slide !== undefined,
		canEdit: canEditOf(state),
		editing: state.editingActive && !state.editor.masterViewTarget,
		canInsert: !state.protectedViewActive && !state.readOnlyRec.locked,
		hasSelection: element !== null,
		selectionHasHyperlink: Boolean(element?.actionClick),
		selectionHasEquation: selectionEquation(element) !== null,
		layoutCount: scopedLayoutsOf(state, source.layouts()).length,
		hasCurrentLayout: resetSlideLayoutPath(slide) !== undefined,
		hasHandler: state.loader.handler !== null,
	};
}

/** The shared find port over the viewer's own `FindReplaceState`. */
export function findReplacePort(find: FindReplaceState): BuildingBlockFindReplacePort {
	return {
		open: find.open,
		setOpen: (open) => {
			if (open !== find.open) {
				find.toggle();
			}
		},
		query: find.query,
		replacement: find.replacement,
		options: {
			matchCase: find.matchCase,
			includeNotes: find.includeNotes,
			includeChartTitles: find.includeChartTitles,
			includeChartData: find.includeChartData,
		},
		matches: find.results,
		currentIndex: find.index,
		setQuery: (query) => {
			find.query = query;
		},
		setReplacement: (replacement) => {
			find.replacement = replacement;
		},
		setOptions: (change) => {
			if (change.matchCase !== undefined) {
				find.matchCase = change.matchCase;
			}
			if (change.includeNotes !== undefined) {
				find.includeNotes = change.includeNotes;
			}
			if (change.includeChartTitles !== undefined) {
				find.includeChartTitles = change.includeChartTitles;
			}
			if (change.includeChartData !== undefined) {
				find.includeChartData = change.includeChartData;
			}
		},
		find: () => find.search(),
		navigate: (delta) => (delta > 0 ? find.next() : find.prev()),
		replaceCurrent: () => find.replaceCurrent(),
		replaceAll: () => find.replaceAll(),
	};
}

/** The shared session's port over the Svelte viewer state, read at call time. */
export function createSveltePort(source: SveltePortSource): BuildingBlockSessionPort {
	const { state } = source;
	const { editor } = state;
	const patchSelection = (patch: Partial<PptxElement>) => {
		const element = selectedElementOf(state);
		if (!element) {
			return buildingBlockFailure('no-selection');
		}
		editor.applyElementPatch(element.id, patch);
		return BUILDING_BLOCK_OK;
	};
	return {
		availabilityInput: () => svelteAvailabilityInput(source),
		onCancel: () => {
			state.parityUi.shortcutsOpen = false;
		},
		fileLimits: () => source.options().fileLimits,
		onRequest: (request) => source.options().onRequest?.(request),
		target: () => ({
			deck: state.loader.handler,
			slideId: editor.slides[state.viewer.current]?.id,
			slideIndex: state.viewer.current,
		}),
		canvasSize: () => state.loader.canvasSize,
		selectedElement: () => selectedElementOf(state),
		slideCount: () => editor.slides.length,
		activeSlide: () => editor.slides[state.viewer.current],
		scopedLayouts: () => scopedLayoutsOf(state, source.layouts()),
		addElement: (element) => {
			editor.insertElement(element);
		},
		insertMediaFile: async (file, sameTarget) => {
			const element = await buildMediaInsertElement(file, state.loader.canvasSize);
			if (!element) {
				return buildingBlockFailure('decode-failed');
			}
			if (!sameTarget()) {
				return buildingBlockFailure('stale-target');
			}
			editor.insertElement(element);
			return BUILDING_BLOCK_OK;
		},
		applyHyperlink: (data) => {
			const patch = hyperlinkEditPatch(data);
			return patch.ok ? patchSelection({ actionClick: patch.actionClick }) : patch;
		},
		removeHyperlink: () => {
			const element = selectedElementOf(state);
			if (element && !element.actionClick) {
				return buildingBlockFailure('no-hyperlink');
			}
			return patchSelection({ actionClick: undefined });
		},
		insertSmartArt: (layout, items) => {
			editor.insertElement(buildSmartArtInsertElement(layout, items, state.loader.canvasSize));
		},
		insertChart: (kind, options) => {
			editor.insertElement(buildChartInsertElement(kind, state.loader.canvasSize, options));
		},
		insertEquation: (omml) => {
			editor.insertElement(buildEquationInsertElement(omml, state.loader.canvasSize));
		},
		updateEquation: (omml) => {
			const element = selectedElementOf(state);
			if (element) {
				// The equation controller replaces the OMML of the element it is editing.
				editor.equationOps.editingId = element.id;
				editor.equationOps.apply(omml);
			}
		},
		findReplace: () => findReplacePort(state.findReplace),
		applyLayout: async (path) => {
			// A layout the handler cannot resolve rejects, so the session reports `unknown-layout`.
			if ((await editor.slidesOps.applyLayout(path)) === null) {
				throw new Error(`Layout ${path} was not applied`);
			}
		},
	};
}
