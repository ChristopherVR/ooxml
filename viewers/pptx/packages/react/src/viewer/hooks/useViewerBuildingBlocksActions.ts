import {
	createPptxBuildingBlockSession,
	registerBuildingBlockGroups,
	resetSlideLayoutPath,
	selectionEquation,
} from 'ooxml-ui/pptx';
import type {
	BuildingBlockAvailabilityInput,
	BuildingBlockSessionPort,
	PptxBuildingBlockSession,
	ReadOnlyRecommendation,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide } from 'pptx-viewer-core';
import { useEffect, useMemo, useRef } from 'react';

import { useController } from './useController';
import type { EditorHistoryResult } from './useEditorHistory';
import type { EditorOperationsResult } from './useEditorOperations';
import { useLayoutSwitching } from './useLayoutSwitching';
import type { UsePresentationModeResult } from './usePresentationMode';
import type { PropertyHandlersResult } from './usePropertyHandlers';
import { scopeLayoutOptionsToActiveSlide } from './useScopedLayoutOptions';
import type {
	UseViewerBuildingBlocksActionsOptions,
	ViewerBuildingBlockActions,
} from './useViewerBuildingBlocks-actions-types';
import { useFileDialogPorts } from './useViewerBuildingBlocks-file-dialog-actions';
import { findReplacePort } from './useViewerBuildingBlocks-find-layout-actions';
import { useShowDialogPorts } from './useViewerBuildingBlocks-show-dialog-actions';
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
	propertyHandlers: PropertyHandlersResult;
	content: ArrayBuffer | Uint8Array | null | undefined;
	readOnlyRecommendation: Pick<ReadOnlyRecommendation, 'kind'>;
	hasModifyVerifier: boolean;
	presentation: UsePresentationModeResult;
}

export interface ViewerBuildingBlocksActionsResult {
	actions: ViewerBuildingBlockActions;
	/**
	 * The shared session, for hooks that register further action groups
	 * (`session.register(group, port)`); read `session.actions()` after they ran.
	 */
	session: PptxBuildingBlockSession;
	/** Raise an image / media picker request (what the toolbar's Picture and Video controls call). */
	requestImagePicker: () => void;
	requestMediaPicker: () => void;
	/** Home > Layout and Reset, wired to the same operations `actions.layout` exposes. */
	applyLayout: (path: string) => void;
	resetSlide: () => void;
	loadLayoutPreviews: ReturnType<typeof useLayoutSwitching>['loadLayoutPreviews'];
}

type Latest = UseViewerBuildingBlocksActionsInput & {
	availabilityInput: BuildingBlockAvailabilityInput;
	scopedLayouts: ReturnType<typeof scopeLayoutOptionsToActiveSlide>;
	layoutSwitching: ReturnType<typeof useLayoutSwitching>;
	selectedElement: PptxElement | null;
};

/**
 * The host-owned workflows of the building-block API, as the React store adapter of the shared
 * session (`createPptxBuildingBlockSession` in `ooxml-ui/pptx`): this hook gathers the
 * availability input and the editor operations into the session's port, registers the built-in
 * dialog groups and reads the pending request through `useController`. Every operation runs the
 * viewer's existing code, so selection, history and dirty notifications behave as for the bundled
 * viewer, and is refused with a code when the viewer is read-only, protected or not ready.
 */
export function useViewerBuildingBlocksActions(
	input: UseViewerBuildingBlocksActionsInput,
): ViewerBuildingBlocksActionsResult {
	const { state, dialogs, activeSlide, handlerRef } = input;
	// The first of the effective selection (a plain click selects through the id set), as the insert handlers read it.
	const selectedElement: PptxElement | null = state.selectedElements[0] ?? null;

	const scopedLayouts = useMemo(
		() => scopeLayoutOptionsToActiveSlide(state.layoutOptions, activeSlide),
		[state.layoutOptions, activeSlide],
	);
	const layoutSwitching = useLayoutSwitching({
		handler: handlerRef.current,
		slides: input.slides,
		activeSlideIndex: input.activeSlideIndex,
		ops: input.editorOps.ops,
		history: input.history,
		onTemplateElementsChanged: (slideId, elements) =>
			state.setTemplateElementsBySlideId((prev) => ({ ...prev, [slideId]: elements })),
	});

	const availabilityInput: BuildingBlockAvailabilityInput = {
		ready: !state.loading && !state.error,
		hasActiveSlide: Boolean(activeSlide),
		canEdit: input.canEdit,
		editing: input.mode === 'edit',
		canInsert: input.canInsert,
		hasSelection: selectedElement !== null,
		selectionHasHyperlink: Boolean(selectedElement?.actionClick),
		selectionHasEquation: selectionEquation(selectedElement) !== null,
		layoutCount: scopedLayouts.length,
		hasCurrentLayout: resetSlideLayoutPath(activeSlide) !== undefined,
		hasHandler: handlerRef.current !== null,
	};

	// Operations run long after the render that made them: the port reads the latest of everything.
	const next = { ...input, availabilityInput, scopedLayouts, layoutSwitching, selectedElement };
	const latest = useRef<Latest>(next);
	latest.current = next;
	const sessionRef = useRef<PptxBuildingBlockSession | null>(null);
	sessionRef.current ??= createPptxBuildingBlockSession(createReactPort(latest));
	const session = sessionRef.current;
	// Subscribes this hook to the pending request; `session.actions()` reads it below.
	useController(session.requests);

	const fileDialogPorts = useFileDialogPorts({
		state,
		dialogs,
		propertyHandlers: input.propertyHandlers,
		history: input.history,
		content: input.content,
		recommendation: input.readOnlyRecommendation,
		hasModifyVerifier: input.hasModifyVerifier,
		request: (request) => void session.request(request),
	});
	const showDialogPorts = useShowDialogPorts({
		state,
		dialogs,
		presentation: input.presentation,
		slides: input.slides,
		session,
	});
	// The built-in dialog groups, registered once for the session's lifetime (their ports are stable).
	const registered = useRef(false);
	if (!registered.current) {
		registered.current = true;
		registerBuildingBlockGroups(session, { ...fileDialogPorts, ...showDialogPorts });
	}

	// The toolbar and the canvas (double-clicking an equation, a link's context menu entry) open the
	// internal dialogs by flag; nothing renders them here, so each flag becomes a host request.
	const { isHyperlinkDialogOpen, isSmartArtDialogOpen, isEquationDialogOpen } = dialogs;
	useEffect(() => {
		if (isHyperlinkDialogOpen) {
			dialogs.setIsHyperlinkDialogOpen(false);
			session.request({ kind: 'hyperlink' });
		}
	}, [isHyperlinkDialogOpen, dialogs, session]);
	useEffect(() => {
		if (isSmartArtDialogOpen) {
			dialogs.setIsSmartArtDialogOpen(false);
			session.request({ kind: 'smartArt' });
		}
	}, [isSmartArtDialogOpen, dialogs, session]);
	useEffect(() => {
		if (isEquationDialogOpen) {
			const existingOmml = dialogs.editingEquationOmml;
			dialogs.setIsEquationDialogOpen(false);
			dialogs.setEditingEquationOmml(null);
			session.request({ kind: 'equation', existingOmml });
		}
	}, [isEquationDialogOpen, dialogs, session]);
	const { isChartDialogOpen } = dialogs;
	useEffect(() => {
		if (isChartDialogOpen) {
			dialogs.setIsChartDialogOpen(false);
			session.request({ kind: 'chart' });
		}
	}, [isChartDialogOpen, dialogs, session]);

	return {
		actions: session.actions(),
		session,
		requestImagePicker: () => void session.request({ kind: 'image' }),
		requestMediaPicker: () => void session.request({ kind: 'media' }),
		applyLayout: (path) => void session.applyLayout(path),
		resetSlide: () => void session.resetSlide(),
		loadLayoutPreviews: layoutSwitching.loadLayoutPreviews,
	};
}

/** The shared session's port over this binding's editor, read through `latest` at call time. */
function createReactPort(latest: { readonly current: Latest }): BuildingBlockSessionPort {
	const insert = () => latest.current.editorOps.insertHandlers;
	return {
		availabilityInput: () => latest.current.availabilityInput,
		fileLimits: () => latest.current.fileLimits,
		onRequest: (request) => latest.current.onRequest?.(request),
		onCancel: () => latest.current.state.setIsShortcutHelpOpen(false),
		target: () => ({
			deck: latest.current.handlerRef.current,
			slideId: latest.current.activeSlide?.id,
			slideIndex: latest.current.activeSlideIndex,
		}),
		canvasSize: () => latest.current.state.canvasSize,
		selectedElement: () => latest.current.selectedElement,
		slideCount: () => latest.current.slides.length,
		activeSlide: () => latest.current.activeSlide,
		scopedLayouts: () => latest.current.scopedLayouts,
		addElement: (element) => insert().addElement(element),
		insertMediaFile: (file, sameTarget) => insert().insertMediaFile(file, sameTarget),
		applyHyperlink: (data) => insert().applyHyperlink(data),
		removeHyperlink: () => insert().removeHyperlink(),
		insertSmartArt: (layout, items) => insert().handleInsertSmartArt(layout, items),
		insertChart: (kind, options) => insert().handleAddChart(kind, options),
		insertEquation: (omml) => insert().handleInsertEquation(omml),
		updateEquation: (omml) => insert().handleUpdateEquation(omml),
		findReplace: () => findReplacePort(latest.current.editorOps.findReplace),
		applyLayout: (path) => latest.current.layoutSwitching.applyLayout(path),
	};
}
