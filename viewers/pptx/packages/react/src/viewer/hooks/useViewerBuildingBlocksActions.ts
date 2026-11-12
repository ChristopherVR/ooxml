import {
	BUILDING_BLOCK_OK,
	BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS,
	buildingBlockFailure,
	createImageElementFromFile,
	describeBuildingBlockActions,
	resetSlideLayoutPath,
	validateHostImageFile,
	validateHostMediaFile,
} from 'ooxml-ui/pptx';
import type { BuildingBlockActionId, BuildingBlockResult, FindReplaceOptions } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

import type { HyperlinkEditData } from '../components/hyperlink-edit-types';
import {
	hyperlinkEditDataFromAction,
	resolveHyperlinkEditResult,
} from '../components/hyperlink-edit-utils';
import type { EditorHistoryResult } from './useEditorHistory';
import type { EditorOperationsResult } from './useEditorOperations';
import { useLayoutSwitching } from './useLayoutSwitching';
import { scopeLayoutOptionsToActiveSlide } from './useScopedLayoutOptions';
import type {
	BuildingBlockRequest,
	UseViewerBuildingBlocksActionsOptions,
	ViewerBuildingBlockActions,
} from './useViewerBuildingBlocks-actions-types';
import type { ViewerDialogsResult } from './useViewerDialogs';
import type { ViewerState } from './useViewerState';

export interface UseViewerBuildingBlocksActionsInput extends UseViewerBuildingBlocksActionsOptions {
	state: ViewerState;
	canEdit: boolean;
	/** False while read-only recommended or in Protected View (inserting is blocked). */
	canInsert: boolean;
	mode: string;
	slides: PptxSlide[];
	activeSlide: PptxSlide | undefined;
	activeSlideIndex: number;
	dialogs: ViewerDialogsResult;
	editorOps: EditorOperationsResult;
	history: EditorHistoryResult;
	handlerRef: { readonly current: PptxHandler | null };
}

export interface ViewerBuildingBlocksActionsResult {
	actions: ViewerBuildingBlockActions;
	/** Raise an image / media picker request (what the toolbar's Picture and Video controls call). */
	requestImagePicker: () => void;
	requestMediaPicker: () => void;
	/** Home > Layout and Reset, wired to the same operations `actions.layout` exposes. */
	applyLayout: (path: string) => void;
	resetSlide: () => void;
	loadLayoutPreviews: ReturnType<typeof useLayoutSwitching>['loadLayoutPreviews'];
}

function equationOf(element: PptxElement | null): Record<string, unknown> | null {
	const segments = (element as { textSegments?: Array<{ equationXml?: unknown }> } | null)
		?.textSegments;
	const found = segments?.find((segment) => segment.equationXml)?.equationXml;
	return found ? (found as Record<string, unknown>) : null;
}

/**
 * The host-owned workflows of the building-block API: operations a custom shell calls from its own
 * pickers and dialogs, plus the requests that tell it when to show them. Every operation runs the
 * viewer's existing insert, find / replace and layout code, so selection, history and dirty
 * notifications behave as for the bundled viewer, and is refused with a code when the viewer is
 * read-only, protected or not ready.
 */
export function useViewerBuildingBlocksActions(
	input: UseViewerBuildingBlocksActionsInput,
): ViewerBuildingBlocksActionsResult {
	const { state, dialogs, editorOps, history, activeSlide, handlerRef } = input;
	// The first of the effective selection (a plain click selects through the id set), as the insert handlers read it.
	const selectedElement: PptxElement | null = state.selectedElements[0] ?? null;
	const { insertHandlers, findReplace } = editorOps;
	const [pending, setPending] = useState<BuildingBlockRequest | null>(null);

	const scopedLayouts = useMemo(
		() => scopeLayoutOptionsToActiveSlide(state.layoutOptions, activeSlide),
		[state.layoutOptions, activeSlide],
	);
	const layoutSwitching = useLayoutSwitching({
		handler: handlerRef.current,
		slides: input.slides,
		activeSlideIndex: input.activeSlideIndex,
		ops: editorOps.ops,
		history,
		onTemplateElementsChanged: (slideId, elements) =>
			state.setTemplateElementsBySlideId((prev) => ({ ...prev, [slideId]: elements })),
	});

	const equation = equationOf(selectedElement);
	const availability = describeBuildingBlockActions({
		ready: !state.loading && !state.error,
		hasActiveSlide: Boolean(activeSlide),
		canEdit: input.canEdit,
		editing: input.mode === 'edit',
		canInsert: input.canInsert,
		hasSelection: selectedElement !== null,
		selectionHasHyperlink: Boolean(selectedElement?.actionClick),
		selectionHasEquation: equation !== null,
		layoutCount: scopedLayouts.length,
		hasCurrentLayout: resetSlideLayoutPath(activeSlide) !== undefined,
		hasHandler: handlerRef.current !== null,
	});

	// Operations run long after the render that made them: they read the latest of everything.
	const latest = useRef({ ...input, availability, scopedLayouts, layoutSwitching });
	latest.current = { ...input, availability, scopedLayouts, layoutSwitching };
	const refuse = (id: BuildingBlockActionId): BuildingBlockResult | undefined => {
		const { available, reason } = latest.current.availability[id];
		return available ? undefined : buildingBlockFailure(reason ?? 'not-ready');
	};

	const request = useCallback((next: BuildingBlockRequest, id: BuildingBlockActionId) => {
		if (!latest.current.availability[id].available) {
			return;
		}
		setPending(next);
		latest.current.onRequest?.(next);
	}, []);
	const complete = (kind: BuildingBlockRequest['kind'], result: BuildingBlockResult) => {
		if (result.ok) {
			setPending((current) => (current?.kind === kind ? null : current));
		}
		return result;
	};

	// The toolbar and the canvas (double-clicking an equation, a link's context menu entry) open the
	// internal dialogs by flag; nothing renders them here, so each flag becomes a host request.
	const { isHyperlinkDialogOpen, isSmartArtDialogOpen, isEquationDialogOpen } = dialogs;
	useEffect(() => {
		if (isHyperlinkDialogOpen) {
			dialogs.setIsHyperlinkDialogOpen(false);
			request({ kind: 'hyperlink' }, 'hyperlink');
		}
	}, [isHyperlinkDialogOpen, dialogs, request]);
	useEffect(() => {
		if (isSmartArtDialogOpen) {
			dialogs.setIsSmartArtDialogOpen(false);
			request({ kind: 'smartArt' }, 'insertSmartArt');
		}
	}, [isSmartArtDialogOpen, dialogs, request]);
	useEffect(() => {
		if (isEquationDialogOpen) {
			const existingOmml = dialogs.editingEquationOmml;
			dialogs.setIsEquationDialogOpen(false);
			dialogs.setEditingEquationOmml(null);
			request({ kind: 'equation', existingOmml }, existingOmml ? 'editEquation' : 'insertEquation');
		}
	}, [isEquationDialogOpen, dialogs, request]);

	const insertImage = useCallback(async (file: File): Promise<BuildingBlockResult> => {
		const denied = refuse('insertImage');
		if (denied) {
			return denied;
		}
		const checked = validateHostImageFile(file, latest.current.fileLimits?.imageBytes);
		if (!checked.ok) {
			return checked;
		}
		const handler = latest.current.handlerRef.current;
		const slideId = latest.current.activeSlide?.id;
		const element = await createImageElementFromFile(file, latest.current.state.canvasSize);
		if (!element) {
			return buildingBlockFailure('decode-failed');
		}
		// The deck or slide may have changed while the picture decoded.
		const now = latest.current;
		if (now.handlerRef.current !== handler || now.activeSlide?.id !== slideId) {
			return buildingBlockFailure('stale-target');
		}
		const stillDenied = refuse('insertImage');
		if (stillDenied) {
			return stillDenied;
		}
		now.editorOps.insertHandlers.addElement(element);
		return complete('image', BUILDING_BLOCK_OK);
		// oxlint-disable-next-line react/memo-dependencies -- reads `latest`
	}, []);

	const insertMedia = useCallback(async (file: File): Promise<BuildingBlockResult> => {
		const denied = refuse('insertMedia');
		if (denied) {
			return denied;
		}
		const checked = validateHostMediaFile(file, latest.current.fileLimits?.mediaBytes);
		if (!checked.ok) {
			return checked;
		}
		return complete('media', await latest.current.editorOps.insertHandlers.insertMediaFile(file));
		// oxlint-disable-next-line react/memo-dependencies -- reads `latest`
	}, []);

	const hyperlinkCurrent = hyperlinkEditDataFromAction(selectedElement?.actionClick);
	const hyperlink = {
		current: hyperlinkCurrent,
		slideCount: input.slides.length,
		apply: (data: HyperlinkEditData): BuildingBlockResult => {
			const denied = refuse('hyperlink');
			if (denied) {
				return denied;
			}
			const resolved = resolveHyperlinkEditResult(data);
			if (!resolved.url && !resolved.action) {
				return buildingBlockFailure('invalid-target');
			}
			return complete('hyperlink', insertHandlers.applyHyperlink(data));
		},
		remove: (): BuildingBlockResult =>
			refuse('removeHyperlink') ?? insertHandlers.removeHyperlink(),
	};

	const gated = (
		id: BuildingBlockActionId,
		kind: BuildingBlockRequest['kind'],
		run: () => void,
	): BuildingBlockResult => {
		const denied = refuse(id);
		if (denied) {
			return denied;
		}
		run();
		return complete(kind, BUILDING_BLOCK_OK);
	};

	const applyLayoutChecked = async (
		path: string,
		requireListed: boolean,
	): Promise<BuildingBlockResult> => {
		const denied = refuse('applyLayout');
		if (denied) {
			return denied;
		}
		if (requireListed && !latest.current.scopedLayouts.some((layout) => layout.path === path)) {
			return buildingBlockFailure('unknown-layout');
		}
		try {
			await latest.current.layoutSwitching.applyLayout(path);
			return BUILDING_BLOCK_OK;
		} catch {
			return buildingBlockFailure('unknown-layout');
		}
	};
	const resetLayout = async (): Promise<BuildingBlockResult> => {
		const path = resetSlideLayoutPath(latest.current.activeSlide);
		if (path === undefined) {
			return buildingBlockFailure('no-layout');
		}
		return refuse('resetLayout') ?? applyLayoutChecked(path, false);
	};

	const matches = findReplace.findResults;
	const currentIndex = findReplace.findResultIndex;
	const options: FindReplaceOptions = {
		matchCase: findReplace.findMatchCase,
		includeNotes: findReplace.findIncludeNotes,
		includeChartTitles: findReplace.findIncludeChartTitles,
		includeChartData: findReplace.findIncludeChartData,
	};

	const actions: ViewerBuildingBlockActions = {
		availability,
		unmountedDialogActions: BUILDING_BLOCK_UNMOUNTED_DIALOG_ACTIONS,
		pending,
		cancel: () => setPending(null),
		insertImage,
		insertMedia,
		hyperlink,
		smartArt: {
			insert: (layout, items) =>
				gated('insertSmartArt', 'smartArt', () =>
					insertHandlers.handleInsertSmartArt(layout, items),
				),
		},
		equation: {
			current: equation,
			insert: (omml) =>
				gated('insertEquation', 'equation', () => insertHandlers.handleInsertEquation(omml)),
			update: (omml) =>
				gated('editEquation', 'equation', () => insertHandlers.handleUpdateEquation(omml)),
		},
		findReplace: {
			open: findReplace.findReplaceOpen,
			setOpen: findReplace.setFindReplaceOpen,
			query: findReplace.findQuery,
			replacement: findReplace.replaceQuery,
			options,
			matches,
			currentIndex,
			currentMatch: matches[currentIndex] ?? null,
			setQuery: findReplace.setFindQuery,
			setReplacement: findReplace.setReplaceQuery,
			setOptions: (change) => {
				if (change.matchCase !== undefined) {
					findReplace.setFindMatchCase(change.matchCase);
				}
				if (change.includeNotes !== undefined) {
					findReplace.setFindIncludeNotes(change.includeNotes);
				}
				if (change.includeChartTitles !== undefined) {
					findReplace.setFindIncludeChartTitles(change.includeChartTitles);
				}
				if (change.includeChartData !== undefined) {
					findReplace.setFindIncludeChartData(change.includeChartData);
				}
			},
			find: findReplace.performFind,
			next: () => findReplace.navigateFindResult(1),
			previous: () => findReplace.navigateFindResult(-1),
			replaceCurrent: () =>
				refuse('replace') ??
				(matches[currentIndex] === undefined
					? buildingBlockFailure('no-match')
					: (findReplace.handleReplace(), BUILDING_BLOCK_OK)),
			replaceAll: () =>
				refuse('replace') ??
				(matches.length === 0
					? buildingBlockFailure('no-match')
					: (findReplace.handleReplaceAll(), BUILDING_BLOCK_OK)),
		},
		layout: {
			current: scopedLayouts.find((layout) => layout.path === activeSlide?.layoutPath),
			currentPath: activeSlide?.layoutPath,
			available: scopedLayouts,
			apply: (path) => applyLayoutChecked(path, true),
			reset: resetLayout,
		},
	};

	return {
		actions,
		requestImagePicker: () => request({ kind: 'image' }, 'insertImage'),
		requestMediaPicker: () => request({ kind: 'media' }, 'insertMedia'),
		applyLayout: (path) => void applyLayoutChecked(path, true),
		resetSlide: () => void resetLayout(),
		loadLayoutPreviews: layoutSwitching.loadLayoutPreviews,
	};
}
