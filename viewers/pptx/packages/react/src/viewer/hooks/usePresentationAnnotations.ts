/**
 * Slide-show ink for the React binding: the shared
 * `createPresentationInkController` (tool, colours, per-slide strokes, eraser,
 * laser, Ctrl+M markup visibility and the popup toolbar's auto-hide) read
 * through `useController`. The hook only follows the active slide and the
 * show's lifetime, and reveals the toolbar on mouse movement.
 */
import {
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
} from 'ooxml-ui/pptx';
import type { PresentationInkController } from 'ooxml-ui/pptx';
import { useCallback, useEffect, useMemo, useState } from 'react';

import { useController } from './useController';
import type {
	AnnotationStroke,
	LaserPosition,
	PresentationTool,
	SlideAnnotationMap,
	UsePresentationAnnotationsInput,
	UsePresentationAnnotationsResult,
} from './usePresentationAnnotations.types';

export type {
	AnnotationStroke,
	LaserPosition,
	PresentationTool,
	SlideAnnotationMap,
	UsePresentationAnnotationsInput,
	UsePresentationAnnotationsResult,
};

export function usePresentationAnnotations(
	input: UsePresentationAnnotationsInput,
): UsePresentationAnnotationsResult {
	const { isActive, activeSlideIndex, popupToolbarEnabled = true } = input;
	const [ink] = useState<PresentationInkController>(() => createPresentationInkController());
	const state = useController(ink);

	useEffect(() => ink.setSlide(activeSlideIndex), [ink, activeSlideIndex]);
	useEffect(() => () => ink.dispose(), [ink]);

	// Leaving the show disarms; the strokes stay for the keep-or-discard prompt.
	useEffect(() => {
		if (!isActive) {
			ink.reset();
			return;
		}
		if (!popupToolbarEnabled) {
			// Options > Advanced > "Show popup toolbar" is off: never auto-show.
			return;
		}
		const reveal = (): void => ink.revealToolbar();
		window.addEventListener('mousemove', reveal);
		return () => window.removeEventListener('mousemove', reveal);
	}, [ink, isActive, popupToolbarEnabled]);

	const whenActive = useCallback(
		(action: (x: number, y: number) => void) =>
			(x: number, y: number): void => {
				if (isActive) {
					action(x, y);
				}
			},
		[isActive],
	);
	const allSlideAnnotations = useMemo<SlideAnnotationMap>(
		() => presentationInkBySlide(state),
		[state],
	);
	const annotationStrokes = useMemo(() => presentationInkOnSlide(state), [state]);

	return {
		presentationTool: state.tool,
		setPresentationTool: useCallback(
			(tool: PresentationTool) => {
				if (isActive) {
					ink.toggleTool(tool);
				}
			},
			[ink, isActive],
		),
		selectPresentationTool: useCallback(
			(tool: PresentationTool) => {
				if (isActive) {
					ink.setTool(tool);
				}
			},
			[ink, isActive],
		),
		penColor: state.penColor,
		setPenColor: ink.setPenColor,
		highlighterColor: state.highlighterColor,
		setHighlighterColor: ink.setHighlighterColor,
		annotationStrokes,
		currentStroke: state.current,
		laserPosition: state.laser,
		toolbarVisible: state.toolbarVisible,
		setToolbarVisible: ink.setToolbarVisible,
		inkMarkupVisible: state.markupVisible,
		setInkMarkupVisible: ink.setMarkupVisible,
		handlePointerDown: useMemo(
			() =>
				whenActive((x, y) => {
					if (state.tool === 'pen' || state.tool === 'highlighter') {
						ink.pointerDown({ x, y });
					}
				}),
			[ink, state.tool, whenActive],
		),
		handlePointerMove: useMemo(
			() =>
				whenActive((x, y) => {
					if (ink.getSnapshot().current) {
						ink.pointerMove({ x, y });
					}
				}),
			[ink, whenActive],
		),
		handlePointerUp: ink.pointerUp,
		handleLaserMove: useMemo(
			() =>
				whenActive((x, y) => {
					if (ink.getSnapshot().tool === 'laser') {
						ink.pointerMove({ x, y });
					}
				}),
			[ink, whenActive],
		),
		handleLaserLeave: ink.pointerLeave,
		clearAnnotations: ink.clearSlide,
		eraseAtPoint: useMemo(() => whenActive((x, y) => ink.eraseAt({ x, y })), [ink, whenActive]),
		allSlideAnnotations,
		hasAnyAnnotations: state.strokes.length > 0,
		clearAllAnnotations: ink.clearAll,
	};
}
