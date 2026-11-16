/**
 * `usePresentationAnnotations`: slide-show ink for the Vue binding.
 *
 * The tool, colours, per-slide strokes, stroke in progress, eraser and laser
 * live in the shared `createPresentationInkController`; this composable reads
 * it through the one store adapter (`useController`), follows the active slide
 * and names the state for the overlay and toolbar templates.
 *
 * @module composables/usePresentationAnnotations
 */

import {
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
} from 'ooxml-ui/pptx';
import type { PresentationAnnotationStroke, PresentationPointerTool } from 'ooxml-ui/pptx';
import { computed, getCurrentScope, onScopeDispose, toValue, watch } from 'vue';
import type { ComputedRef, MaybeRefOrGetter, WritableComputedRef } from 'vue';

import { useController } from './useController';

export {
	PRESENTATION_ERASER_RADIUS as ERASER_RADIUS,
	PRESENTATION_HIGHLIGHTER_OPACITY as HIGHLIGHTER_OPACITY,
	PRESENTATION_HIGHLIGHTER_WIDTH as HIGHLIGHTER_WIDTH,
	PRESENTATION_PEN_WIDTH as PEN_WIDTH,
} from 'ooxml-ui/pptx';

export type PresentationTool = PresentationPointerTool;
export interface AnnotationPoint {
	x: number;
	y: number;
}
export type AnnotationStroke = PresentationAnnotationStroke;
export type LaserPosition = AnnotationPoint;
/** Annotations grouped by slide index. */
export type SlideAnnotationMap = Map<number, AnnotationStroke[]>;

export interface UsePresentationAnnotationsOptions {
	/** Whether presentation mode (and thus annotation capture) is active. */
	isActive: MaybeRefOrGetter<boolean>;
	/** Current slide index, used to track which slide annotations belong to. */
	activeSlideIndex: MaybeRefOrGetter<number>;
}

export interface UsePresentationAnnotationsResult {
	/** The armed tool. Writing it arms without toggling. */
	presentationTool: WritableComputedRef<PresentationTool>;
	/** Arm `tool`, or disarm it when it is already armed. */
	setPresentationTool: (tool: PresentationTool) => void;
	/** Arm `tool` without toggling (pointer options, keyboard, Blackboard). */
	selectPresentationTool: (tool: PresentationTool) => void;
	penColor: ComputedRef<string>;
	setPenColor: (color: string) => void;
	highlighterColor: ComputedRef<string>;
	setHighlighterColor: (color: string) => void;
	/** Completed strokes on the active slide. */
	annotationStrokes: ComputedRef<AnnotationStroke[]>;
	currentStroke: ComputedRef<AnnotationStroke | null>;
	laserPosition: ComputedRef<LaserPosition | null>;
	handlePointerDown: (x: number, y: number) => void;
	handlePointerMove: (x: number, y: number) => void;
	handlePointerUp: () => void;
	handleLaserMove: (x: number, y: number) => void;
	handleLaserLeave: () => void;
	eraseAtPoint: (x: number, y: number) => void;
	/** Clear the active slide's annotations. */
	clearAnnotations: () => void;
	/** Clear annotations on every slide. */
	clearAllAnnotations: () => void;
	/** All annotations across all slides (for persistence on exit). */
	allSlideAnnotations: ComputedRef<SlideAnnotationMap>;
	hasAnyAnnotations: ComputedRef<boolean>;
}

export function usePresentationAnnotations(
	options: UsePresentationAnnotationsOptions,
): UsePresentationAnnotationsResult {
	const ink = createPresentationInkController();
	const state = useController(ink);
	const active = (): boolean => toValue(options.isActive);
	ink.setSlide(toValue(options.activeSlideIndex));
	watch(() => toValue(options.activeSlideIndex), ink.setSlide, { flush: 'sync' });
	watch(
		active,
		(isActive) => {
			if (!isActive) {
				ink.reset();
			}
		},
		{ flush: 'sync' },
	);
	if (getCurrentScope()) {
		onScopeDispose(ink.dispose);
	}
	const whenActive =
		(action: (x: number, y: number) => void) =>
		(x: number, y: number): void => {
			if (active()) {
				action(x, y);
			}
		};
	const laserMove = whenActive((x, y) => {
		if (state.value.tool === 'laser') {
			ink.pointerMove({ x, y });
		}
	});

	return {
		presentationTool: computed({
			get: () => state.value.tool,
			set: (tool) => ink.setTool(tool),
		}),
		setPresentationTool: (tool) => {
			if (active()) {
				ink.toggleTool(tool);
			}
		},
		selectPresentationTool: (tool) => {
			if (active()) {
				ink.setTool(tool);
			}
		},
		penColor: computed(() => state.value.penColor),
		setPenColor: ink.setPenColor,
		highlighterColor: computed(() => state.value.highlighterColor),
		setHighlighterColor: ink.setHighlighterColor,
		annotationStrokes: computed(() => presentationInkOnSlide(state.value)),
		currentStroke: computed(() => state.value.current),
		laserPosition: computed(() => state.value.laser),
		handlePointerDown: whenActive((x, y) => ink.pointerDown({ x, y })),
		handlePointerMove: whenActive((x, y) => ink.pointerMove({ x, y })),
		handlePointerUp: ink.pointerUp,
		handleLaserMove: laserMove,
		handleLaserLeave: ink.pointerLeave,
		eraseAtPoint: whenActive((x, y) => ink.eraseAt({ x, y })),
		clearAnnotations: ink.clearSlide,
		clearAllAnnotations: ink.clearAll,
		allSlideAnnotations: computed(() => presentationInkBySlide(state.value)),
		hasAnyAnnotations: computed(() => state.value.strokes.length > 0),
	};
}
