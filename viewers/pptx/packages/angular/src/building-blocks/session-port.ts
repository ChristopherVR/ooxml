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
import { mediaElementFromFile } from 'ooxml-ui/pptx';
import type { PptxElement, PptxLayoutOption } from 'pptx-viewer-core';

import { newChartElement } from '../viewer/editor-insert';
import type { LoadNoticesService } from '../viewer/load-notices.service';
import { buildSmartArtInsertElement } from '../viewer/smart-art-insert-helpers';
import type { ViewerCollaborationShellService } from '../viewer/viewer-collaboration-shell.service';
import { buildEquationElement, buildEquationSegment } from '../viewer/viewer-extra-dialogs-helpers';
import { insertPublicElement } from '../viewer/viewer-public-insertion';
import type { ViewerBuildingBlocksOptions } from './types';

/** What the ports read: the shell's services plus the blocks' own options and find bar. */
export interface AngularPortSource {
	shell: ViewerCollaborationShellService;
	notices: LoadNoticesService;
	options(): ViewerBuildingBlocksOptions;
	findBar: FindBarController;
	/** The canvas block's element, for committing pending inline text before an insert. */
	canvasHost(): HTMLElement | null;
	onCancel(): void;
}

/** The layouts of the deck's masters, in document order (the same list the ribbon offers). */
function layoutOptions(source: AngularPortSource): PptxLayoutOption[] {
	const options: PptxLayoutOption[] = [];
	const seen = new Set<string>();
	for (const master of source.shell.loader.slideMasters()) {
		for (const layout of master.layouts ?? []) {
			if (layout.path && !seen.has(layout.path)) {
				seen.add(layout.path);
				options.push({
					path: layout.path,
					name: layout.name || layout.path.split('/').pop() || layout.path,
				});
			}
		}
	}
	return options;
}

function selected(source: AngularPortSource): PptxElement | null {
	const { editor, activeSlideIndex } = source.shell;
	const id = editor.selectedIds()[0];
	return id ? (editor.elementById(activeSlideIndex(), id) ?? null) : null;
}

function scoped(source: AngularPortSource): PptxLayoutOption[] {
	return scopeLayoutOptionsToSlide(layoutOptions(source), source.shell.activeSlide()?.layoutPath);
}

/** The availability input of the Angular editor, gathered from the shell's signals. */
export function angularAvailabilityInput(
	source: AngularPortSource,
): BuildingBlockAvailabilityInput {
	const { shell } = source;
	const element = selected(source);
	const slide = shell.activeSlide();
	const canEdit = shell.canEdit();
	return {
		ready: !shell.loader.loading() && !shell.loader.error() && !shell.sourcePending(),
		hasActiveSlide: slide !== undefined,
		canEdit,
		editing: canEdit && !shell.editor.editTemplateMode(),
		canInsert: !source.notices.lockActive(),
		hasSelection: element !== null,
		selectionHasHyperlink: Boolean(element?.actionClick),
		selectionHasEquation: selectionEquation(element) !== null,
		layoutCount: scoped(source).length,
		hasCurrentLayout: resetSlideLayoutPath(slide) !== undefined,
		hasHandler: shell.loader.getHandler() !== undefined,
	};
}

/** The shared session's port over the Angular editor, read at call time. */
export function createAngularPort(source: AngularPortSource): BuildingBlockSessionPort {
	const { shell } = source;
	const { editor } = shell;
	const add = (element: PptxElement): void => {
		insertPublicElement(
			element,
			{
				canEdit: shell.canEdit(),
				mode: 'edit',
				hasActiveSlide: Boolean(shell.activeSlide()),
				editTemplateMode: editor.editTemplateMode(),
			},
			editor,
			shell.activeSlideIndex(),
			() => {
				// Blur synchronously emits the canvas's own text commit before the insert.
				source.canvasHost()?.querySelector<HTMLElement>('[data-inline-editor]')?.blur();
			},
		);
	};
	const patchSelection = (patch: Partial<PptxElement>) => {
		const element = selected(source);
		if (!element) {
			return buildingBlockFailure('no-selection');
		}
		editor.updateElement(shell.activeSlideIndex(), element.id, patch);
		return BUILDING_BLOCK_OK;
	};
	return {
		availabilityInput: () => angularAvailabilityInput(source),
		onCancel: () => source.onCancel(),
		fileLimits: () => source.options().fileLimits,
		onRequest: (request) => source.options().onRequest?.(request),
		target: () => ({
			deck: shell.loader.getHandler() ?? null,
			slideId: shell.activeSlide()?.id,
			slideIndex: shell.activeSlideIndex(),
		}),
		canvasSize: () => shell.loader.canvasSize(),
		selectedElement: () => selected(source),
		slideCount: () => editor.slides().length,
		activeSlide: () => shell.activeSlide(),
		scopedLayouts: () => scoped(source),
		addElement: add,
		insertMediaFile: async (file, sameTarget) => {
			const element = await mediaElementFromFile(file, shell.loader.canvasSize());
			if (!element) {
				return buildingBlockFailure('decode-failed');
			}
			if (!sameTarget()) {
				return buildingBlockFailure('stale-target');
			}
			add(element);
			return BUILDING_BLOCK_OK;
		},
		applyHyperlink: (data) => {
			const patch = hyperlinkEditPatch(data);
			return patch.ok ? patchSelection({ actionClick: patch.actionClick }) : patch;
		},
		removeHyperlink: () => {
			const element = selected(source);
			if (element && !element.actionClick) {
				return buildingBlockFailure('no-hyperlink');
			}
			return patchSelection({ actionClick: undefined });
		},
		insertSmartArt: (layout, items) => add(buildSmartArtInsertElement(layout, items)),
		insertChart: (kind, options) => add(newChartElement(kind, options)),
		insertEquation: (omml) => add(buildEquationElement(omml)),
		updateEquation: (omml) => {
			const element = selected(source);
			if (element) {
				editor.updateElement(shell.activeSlideIndex(), element.id, {
					textSegments: [buildEquationSegment(omml)],
				});
			}
		},
		findReplace: () => source.findBar.port(),
		applyLayout: async (path) => {
			const before = editor.slides();
			await editor.applyLayout(shell.activeSlideIndex(), path);
			// The editor swallows a layout the handler rejects; no new slide list means no apply.
			if (editor.slides() === before) {
				throw new Error(`Layout ${path} was not applied`);
			}
		},
	};
}
