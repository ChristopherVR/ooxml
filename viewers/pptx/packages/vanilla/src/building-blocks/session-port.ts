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
	BuildingBlockSessionPort,
	FindBarController,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxLayoutOption, PptxSlide } from 'pptx-viewer-core';

import { collectLayoutOptions } from '../viewer/editor/editing-chrome-sync';
import { findActiveElement } from '../viewer/editor/editor-active-elements';
import { mediaElementFromFile } from '../viewer/editor/editor-insert-media';
import type { PptxViewer } from '../viewer/PptxViewer';
import type { ViewerState } from '../viewer/state';
import type { ViewerBuildingBlocksOptions } from './types';

/** What the port reads: the mounted viewer (null before the canvas mounts) and the options. */
export interface VanillaPortSource {
	viewer(): PptxViewer | null;
	options(): ViewerBuildingBlocksOptions;
	findBar: FindBarController;
	/** `actions.cancel()` ran: close the shortcut help the host mirrors. */
	onCancel(): void;
}

const NOT_MOUNTED: BuildingBlockAvailabilityInput = {
	ready: false,
	hasActiveSlide: false,
	canEdit: false,
	editing: false,
	canInsert: false,
	hasSelection: false,
	selectionHasHyperlink: false,
	selectionHasEquation: false,
	layoutCount: 0,
	hasCurrentLayout: false,
	hasHandler: false,
};

function activeSlide(state: ViewerState): PptxSlide | undefined {
	return state.slides[state.currentSlide];
}

function selected(state: ViewerState): PptxElement | null {
	const id = state.selectedElementId ?? state.selectedElementIds[0];
	return id ? (findActiveElement(state, id) ?? null) : null;
}

function scopedLayouts(state: ViewerState): PptxLayoutOption[] {
	return scopeLayoutOptionsToSlide(
		collectLayoutOptions(state) as PptxLayoutOption[],
		activeSlide(state)?.layoutPath,
	);
}

/** The availability input of the vanilla viewer, gathered from its store. */
export function vanillaAvailabilityInput(
	viewer: PptxViewer | null,
): BuildingBlockAvailabilityInput {
	if (!viewer) {
		return NOT_MOUNTED;
	}
	const state = viewer.store.get();
	const element = selected(state);
	const slide = activeSlide(state);
	return {
		ready: !state.loading && !state.error,
		hasActiveSlide: slide !== undefined,
		canEdit: state.editable,
		editing: state.editable && !state.presenting && state.masterViewTarget === null,
		canInsert: !state.protectedView,
		hasSelection: element !== null,
		selectionHasHyperlink: Boolean(element?.actionClick),
		selectionHasEquation: selectionEquation(element) !== null,
		layoutCount: scopedLayouts(state).length,
		hasCurrentLayout: resetSlideLayoutPath(slide) !== undefined,
		hasHandler: viewer.getHandler() !== null,
	};
}

/** The shared session's port over the vanilla viewer, read at call time. */
export function createVanillaPort(source: VanillaPortSource): BuildingBlockSessionPort {
	const state = (): ViewerState | null => source.viewer()?.store.get() ?? null;
	const edit = () => source.viewer()?.editor.getEditActions();
	const patchSelection = (patch: Partial<PptxElement>) => {
		const current = state();
		const element = current ? selected(current) : null;
		const viewer = source.viewer();
		if (!element || !viewer) {
			return buildingBlockFailure('no-selection');
		}
		viewer.editor.applyElementPatch(element.id, patch);
		return BUILDING_BLOCK_OK;
	};
	return {
		availabilityInput: () => vanillaAvailabilityInput(source.viewer()),
		onCancel: () => source.onCancel(),
		fileLimits: () => source.options().fileLimits,
		onRequest: (request) => source.options().onRequest?.(request),
		target: () => {
			const current = state();
			return {
				deck: source.viewer()?.getHandler() ?? null,
				slideId: current ? activeSlide(current)?.id : undefined,
				slideIndex: current?.currentSlide ?? -1,
			};
		},
		canvasSize: () => state()?.canvasSize ?? { width: 0, height: 0 },
		selectedElement: () => {
			const current = state();
			return current ? selected(current) : null;
		},
		slideCount: () => state()?.slides.length ?? 0,
		activeSlide: () => {
			const current = state();
			return current ? activeSlide(current) : undefined;
		},
		scopedLayouts: () => {
			const current = state();
			return current ? scopedLayouts(current) : [];
		},
		addElement: (element) => edit()?.insertElement(element),
		insertMediaFile: async (file, sameTarget) => {
			const viewer = source.viewer();
			if (!viewer) {
				return buildingBlockFailure('not-ready');
			}
			const element = await mediaElementFromFile(viewer.doc, file, viewer.store.get().canvasSize);
			if (!element) {
				return buildingBlockFailure('decode-failed');
			}
			if (!sameTarget()) {
				return buildingBlockFailure('stale-target');
			}
			viewer.editor.getEditActions().insertElement(element);
			return BUILDING_BLOCK_OK;
		},
		applyHyperlink: (data) => {
			const patch = hyperlinkEditPatch(data);
			return patch.ok ? patchSelection({ actionClick: patch.actionClick }) : patch;
		},
		removeHyperlink: () => {
			const current = state();
			const element = current ? selected(current) : null;
			if (element && !element.actionClick) {
				return buildingBlockFailure('no-hyperlink');
			}
			return patchSelection({ actionClick: undefined });
		},
		insertSmartArt: (layout, items) => edit()?.insertSmartArt(layout, items),
		insertChart: (kind, options) => edit()?.insertChart(kind, options),
		insertEquation: (omml) => edit()?.insertEquation(omml),
		updateEquation: (omml) => {
			const current = state();
			const element = current ? selected(current) : null;
			if (element) {
				edit()?.updateEquation(element.id, omml);
			}
		},
		findReplace: () => source.findBar.port(),
		applyLayout: async (path) => {
			// A layout the handler cannot resolve rejects, so the session reports `unknown-layout`.
			if (!(await edit()?.applyLayout(path))) {
				throw new Error(`Layout ${path} was not applied`);
			}
		},
	};
}
