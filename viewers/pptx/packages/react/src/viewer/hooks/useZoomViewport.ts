import {
	attachViewportNavigation,
	clampZoomScale,
	createViewerZoomStore,
	resetViewportScroll,
} from 'ooxml-ui/pptx';
import type { ViewerZoomState, ViewerZoomStore, ViewportFitOptions } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { useRef, useCallback, useEffect } from 'react';

import { MIN_ELEMENT_SIZE, ZOOM_TO_SELECTION_PADDING } from '../constants';
import type { CanvasSize } from '../types';
import { useViewerStore } from './useViewerStore';
import { useViewportFit } from './useViewportFit';
import { useViewportNode } from './useViewportNode';

/** Module scope so the selector identity is stable across renders. */
const selectZoom = (state: ViewerZoomState): number => state.zoom;

/** Axis-aligned bounding box used by zoom / viewport helpers. */
interface SelectionBounds {
	minX: number;
	minY: number;
	maxX: number;
	maxY: number;
}

interface UseZoomViewportInput extends ViewportFitOptions {
	canvasSize: CanvasSize;
	selectedElements: PptxElement[];
	/** True when a plain drag on the slide has nothing to select and should pan (read-only view). */
	dragPans?: boolean;
	/** Id of the slide on the canvas. Showing another slide keeps the zoom and scrolls to its top-left. */
	activeSlideId?: string;
}

export interface UseZoomViewportResult {
	// Refs
	editWrapperRef: React.RefObject<HTMLDivElement | null>;
	canvasStageRef: React.RefObject<HTMLDivElement | null>;
	canvasViewportRef: React.RefObject<HTMLDivElement | null>;
	/** Bind to a custom viewport's ref to track mounts independently of the hook owner. */
	setCanvasViewportNode?: React.RefCallback<HTMLDivElement>;
	renderScaleRef: React.MutableRefObject<number>;
	// State
	scale: number;
	setScale: (scale: number) => void;
	editorDimensions: CanvasSize | null;
	setEditorDimensions: (dims: CanvasSize | null) => void;
	// Derived
	fitScale: number;
	editorScale: number;
	// Actions
	handleZoomIn: () => void;
	handleZoomOut: () => void;
	handleResetZoom: () => void;
	handleZoomToFit: () => void;
	handleZoomToSelection: () => void;
	centerBoundsInViewport: (bounds: SelectionBounds, nextScale: number) => void;
	getCanvasPointFromClient: (clientX: number, clientY: number) => { x: number; y: number } | null;
}

export function useZoomViewport({
	canvasSize,
	selectedElements,
	fitPadding,
	maxFitScale,
	dragPans = false,
	activeSlideId,
}: UseZoomViewportInput): UseZoomViewportResult {
	// ── Refs ──────────────────────────────────────────────────────────────
	const editWrapperRef = useRef<HTMLDivElement>(null);
	const canvasStageRef = useRef<HTMLDivElement>(null);
	const { viewportRef: canvasViewportRef, viewportNode, setCanvasViewportNode } = useViewportNode();
	const renderScaleRef = useRef(1);

	// ── State ─────────────────────────────────────────────────────────────
	// The user's zoom factor lives in the shared zoom store, so the zoom MODEL
	// (not merely the step size) is one definition across all five bindings.
	// Reading it through a selector subscription also means this hook re-renders
	// for a zoom change and not for anything else in the store (issue #145).
	const storeRef = useRef<ViewerZoomStore | null>(null);
	storeRef.current ??= createViewerZoomStore();
	const store = storeRef.current;
	const scale = useViewerStore(store, selectZoom);
	const setScale = useCallback(
		(next: number) => store.dispatch({ type: 'set-zoom', zoom: next }),
		[store],
	);
	const { editorDimensions, setEditorDimensions, effectiveEditorDimensions, fitScale } =
		useViewportFit(canvasSize, canvasViewportRef, { fitPadding, maxFitScale }, viewportNode);

	const editorScale = fitScale * scale;

	// Keep the mutable ref in sync so imperative code has the latest value.
	useEffect(() => {
		renderScaleRef.current = editorScale;
	}, [editorScale]);

	// Publish the measured viewport fit into the store, so consumers can select
	// the effective on-screen scale from one place. A repeated measurement is
	// dropped by the reducer, so a ResizeObserver settling costs nothing.
	useEffect(() => {
		store.dispatch({ type: 'set-fit-scale', fitScale });
	}, [store, fitScale]);

	// ── Actions ───────────────────────────────────────────────────────────

	const centerBoundsInViewport = useCallback(
		(bounds: SelectionBounds, nextScale: number) => {
			const wrapper = editWrapperRef.current;
			const canvasViewport = canvasViewportRef.current;
			if (!wrapper || !canvasViewport) {
				return;
			}

			const boundedScale = clampZoomScale(nextScale);
			const nextEditorScale = fitScale * boundedScale;
			const centerX = (bounds.minX + bounds.maxX) / 2;
			const centerY = (bounds.minY + bounds.maxY) / 2;
			const targetScrollLeft =
				canvasViewport.offsetLeft + centerX * nextEditorScale - wrapper.clientWidth / 2;
			const targetScrollTop =
				canvasViewport.offsetTop + centerY * nextEditorScale - wrapper.clientHeight / 2;

			wrapper.scrollTo({
				left: Math.max(targetScrollLeft, 0),
				top: Math.max(targetScrollTop, 0),
				behavior: 'smooth',
			});
		},
		[canvasViewportRef, fitScale],
	);

	// The shared zoom store keeps each step identical across bindings.
	const handleZoomIn = useCallback(() => {
		store.dispatch({ type: 'zoom-in' });
	}, [store]);

	const handleZoomOut = useCallback(() => {
		store.dispatch({ type: 'zoom-out' });
	}, [store]);

	const handleResetZoom = useCallback(() => {
		store.dispatch({ type: 'zoom-to-fit' });
	}, [store]);

	const handleZoomToFit = useCallback(() => {
		store.dispatch({ type: 'zoom-to-fit' });
	}, [store]);

	const handleZoomToSelection = useCallback(() => {
		if (selectedElements.length === 0) {
			return;
		}

		let minX = Number.POSITIVE_INFINITY;
		let minY = Number.POSITIVE_INFINITY;
		let maxX = Number.NEGATIVE_INFINITY;
		let maxY = Number.NEGATIVE_INFINITY;

		selectedElements.forEach((element) => {
			minX = Math.min(minX, element.x);
			minY = Math.min(minY, element.y);
			maxX = Math.max(maxX, element.x + Math.max(element.width, MIN_ELEMENT_SIZE));
			maxY = Math.max(maxY, element.y + Math.max(element.height, MIN_ELEMENT_SIZE));
		});

		if (
			!Number.isFinite(minX) ||
			!Number.isFinite(minY) ||
			!Number.isFinite(maxX) ||
			!Number.isFinite(maxY)
		) {
			return;
		}

		const selectionBounds: SelectionBounds = { minX, minY, maxX, maxY };

		const boundsWidth = Math.max(selectionBounds.maxX - selectionBounds.minX, MIN_ELEMENT_SIZE);
		const boundsHeight = Math.max(selectionBounds.maxY - selectionBounds.minY, MIN_ELEMENT_SIZE);
		const availableWidth = Math.max(
			effectiveEditorDimensions.width - ZOOM_TO_SELECTION_PADDING,
			MIN_ELEMENT_SIZE,
		);
		const availableHeight = Math.max(
			effectiveEditorDimensions.height - ZOOM_TO_SELECTION_PADDING,
			MIN_ELEMENT_SIZE,
		);
		const targetEditorScale = Math.min(
			availableWidth / boundsWidth,
			availableHeight / boundsHeight,
		);
		const safeFitScale = fitScale > Number.EPSILON ? fitScale : Number.EPSILON;
		const nextScale = clampZoomScale(targetEditorScale / safeFitScale);

		setScale(nextScale);

		window.requestAnimationFrame(() => {
			window.requestAnimationFrame(() => {
				centerBoundsInViewport(selectionBounds, nextScale);
			});
		});
	}, [
		centerBoundsInViewport,
		effectiveEditorDimensions.height,
		effectiveEditorDimensions.width,
		fitScale,
		selectedElements,
		setScale,
	]);

	// Read live by the viewport controller, so a mode change does not re-attach it.
	const dragPansRef = useRef(dragPans);
	useEffect(() => {
		dragPansRef.current = dragPans;
	}, [dragPans]);

	// Ctrl+wheel and pinch zoom at the pointer, and drag to pan, in the
	// viewport. Shared with the other bindings so the gestures match.
	//
	// Follow remounted canvases; retain initial frame retries for object-ref users.
	useEffect(() => {
		if (viewportNode === null) {
			return;
		}
		let raf = 0;
		let detach: (() => void) | null = null;
		const attach = (): void => {
			const viewport = viewportNode ?? canvasViewportRef.current;
			if (!viewport) {
				raf = requestAnimationFrame(attach);
				return;
			}
			detach = attachViewportNavigation(viewport, {
				getZoom: () => store.getState().zoom,
				setZoom: (zoom) => store.dispatch({ type: 'set-zoom', zoom }),
				dragPans: () => dragPansRef.current,
			});
		};
		attach();
		return () => {
			cancelAnimationFrame(raf);
			detach?.();
		};
	}, [canvasViewportRef, viewportNode, store]);

	// The zoom belongs to the window and is kept across slides; the scroll
	// position belongs to the slide it was chosen on (shared policy).
	useEffect(() => {
		const viewport = viewportNode ?? canvasViewportRef.current;
		if (viewport) {
			resetViewportScroll(viewport);
		}
		// Only a slide change resets the scroll, not a remounted viewport.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [activeSlideId]);

	const getCanvasPointFromClient = useCallback(
		(clientX: number, clientY: number): { x: number; y: number } | null => {
			const canvasStage = canvasStageRef.current;
			if (!canvasStage) {
				return null;
			}

			const rect = canvasStage.getBoundingClientRect();
			if (rect.width <= 0 || rect.height <= 0) {
				return null;
			}
			const x = (clientX - rect.left) / editorScale;
			const y = (clientY - rect.top) / editorScale;
			return {
				x: Math.max(0, Math.min(canvasSize.width, x)),
				y: Math.max(0, Math.min(canvasSize.height, y)),
			};
		},
		[canvasSize.height, canvasSize.width, editorScale],
	);

	// ── Return ────────────────────────────────────────────────────────────
	return {
		editWrapperRef,
		canvasStageRef,
		canvasViewportRef,
		setCanvasViewportNode,
		renderScaleRef,
		scale,
		setScale,
		editorDimensions,
		setEditorDimensions,
		fitScale,
		editorScale,
		handleZoomIn,
		handleZoomOut,
		handleResetZoom,
		handleZoomToFit,
		handleZoomToSelection,
		centerBoundsInViewport,
		getCanvasPointFromClient,
	};
}
