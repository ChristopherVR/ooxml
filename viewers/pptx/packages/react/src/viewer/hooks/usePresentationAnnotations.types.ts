/**
 * Types for presentation annotation state.
 */
import type { PresentationAnnotationStroke, PresentationPointerTool } from 'ooxml-ui/pptx';

export type PresentationTool = PresentationPointerTool;

/** One slide-show stroke (the shared ink controller's). */
export type AnnotationStroke = PresentationAnnotationStroke;

export interface LaserPosition {
	x: number;
	y: number;
}

/** Annotations grouped by slide index. */
export type SlideAnnotationMap = Map<number, AnnotationStroke[]>;

export interface UsePresentationAnnotationsInput {
	isActive: boolean;
	/** Current slide index: used to track which slide annotations belong to. */
	activeSlideIndex: number;
	/**
	 * Options > Advanced > "Show popup toolbar" (default true). When false,
	 * mouse movement never auto-reveals the presentation toolbar; it can
	 * still be summoned explicitly (Ctrl+M).
	 */
	popupToolbarEnabled?: boolean;
}

export interface UsePresentationAnnotationsResult {
	presentationTool: PresentationTool;
	/** The toolbar's tool button: arm `tool`, or disarm it when it is already armed. */
	setPresentationTool: (tool: PresentationTool) => void;
	/** Arm `tool` without toggling (pointer options, colour picks, Blackboard). */
	selectPresentationTool: (tool: PresentationTool) => void;
	penColor: string;
	setPenColor: (color: string) => void;
	highlighterColor: string;
	setHighlighterColor: (color: string) => void;
	annotationStrokes: AnnotationStroke[];
	currentStroke: AnnotationStroke | null;
	laserPosition: LaserPosition | null;
	toolbarVisible: boolean;
	setToolbarVisible: (visible: boolean) => void;
	/**
	 * Whether ink markup is drawn (PowerPoint's Ctrl+M). Hiding it keeps the
	 * strokes so they can be shown again, or kept on exit.
	 */
	inkMarkupVisible: boolean;
	setInkMarkupVisible: (visible: boolean) => void;
	handlePointerDown: (x: number, y: number) => void;
	handlePointerMove: (x: number, y: number) => void;
	handlePointerUp: () => void;
	handleLaserMove: (x: number, y: number) => void;
	handleLaserLeave: () => void;
	clearAnnotations: () => void;
	eraseAtPoint: (x: number, y: number) => void;
	/** All annotations across all slides (for persistence on exit). */
	allSlideAnnotations: SlideAnnotationMap;
	/** Whether any annotations exist across all slides. */
	hasAnyAnnotations: boolean;
	/** Clear all annotations across all slides. */
	clearAllAnnotations: () => void;
}
