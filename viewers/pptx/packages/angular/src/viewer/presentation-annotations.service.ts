/**
 * presentation-annotations.service.ts: slide-show ink for the Angular binding,
 * an adapter over the shared `createPresentationInkController` (tool, colours,
 * per-slide strokes, eraser, laser, popup toolbar auto-hide), read as signals.
 *
 * Provide at the component level so lifetime tracks the host overlay:
 * `@Component({ providers: [PresentationAnnotationsService] })`.
 */

import { DestroyRef, Injectable, computed, inject } from '@angular/core';
import {
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
} from 'ooxml-ui/pptx';

import { controllerSignal } from './controller-signal';
import type {
	AnnotationPoint,
	AnnotationStroke,
	LaserPosition,
	PresentationTool,
	SlideAnnotationMap,
} from './presentation-annotations-helpers';

// Re-export types so consumers only need to import from this service.
export type {
	AnnotationPoint,
	AnnotationStroke,
	LaserPosition,
	PresentationTool,
	SlideAnnotationMap,
};

@Injectable()
export class PresentationAnnotationsService {
	private readonly ink = createPresentationInkController();
	private readonly state = controllerSignal(this.ink);

	readonly tool = computed<PresentationTool>(() => this.state().tool);
	readonly penColor = computed(() => this.state().penColor);
	readonly highlighterColor = computed(() => this.state().highlighterColor);
	/** Committed strokes on the active slide (does not include the live stroke). */
	readonly annotationStrokes = computed<readonly AnnotationStroke[]>(() =>
		presentationInkOnSlide(this.state()),
	);
	readonly currentStroke = computed<AnnotationStroke | null>(() => this.state().current);
	readonly laserPosition = computed<LaserPosition | null>(() => this.state().laser);
	readonly toolbarVisible = computed(() => this.state().toolbarVisible);
	/** Any ink on any slide (the keep-or-discard prompt on exit). */
	readonly hasAnyAnnotations = computed(() => this.state().strokes.length > 0);

	constructor() {
		inject(DestroyRef).onDestroy(() => this.ink.dispose());
	}

	/** The toolbar's tool button: arm `tool`, or disarm it when it is already armed. */
	setTool(tool: PresentationTool): void {
		this.ink.toggleTool(tool);
	}

	/** Arm `tool` without toggling (pointer options, colour picks, Blackboard). */
	selectTool(tool: PresentationTool): void {
		this.ink.setTool(tool);
	}

	setPenColor(color: string): void {
		this.ink.setPenColor(color);
	}

	setHighlighterColor(color: string): void {
		this.ink.setHighlighterColor(color);
	}

	setToolbarVisible(visible: boolean): void {
		this.ink.setToolbarVisible(visible);
	}

	/** The show moved to `index`. */
	setActiveSlide(index: number): void {
		this.ink.setSlide(index);
	}

	beginStroke(x: number, y: number): void {
		const tool = this.state().tool;
		if (tool === 'pen' || tool === 'highlighter') {
			this.ink.pointerDown({ x, y });
		}
	}

	extendStroke(x: number, y: number): void {
		if (this.state().current) {
			this.ink.pointerMove({ x, y });
		}
	}

	endStroke(): void {
		this.ink.pointerUp();
	}

	beginErase(x: number, y: number): void {
		if (this.state().tool === 'eraser') {
			this.ink.pointerDown({ x, y });
		}
	}

	continueErase(x: number, y: number): void {
		if (this.state().tool === 'eraser') {
			this.ink.pointerMove({ x, y });
		}
	}

	endErase(): void {
		this.ink.pointerUp();
	}

	moveLaser(x: number, y: number): void {
		if (this.state().tool === 'laser') {
			this.ink.pointerMove({ x, y });
		}
	}

	hideLaser(): void {
		this.ink.pointerLeave();
	}

	/** "Erase All Ink on Slide": the active slide. */
	clearAnnotations(): void {
		this.ink.clearSlide();
	}

	clearAllAnnotations(): void {
		this.ink.clearAll();
	}

	/** Every slide's strokes (the "keep as ink" handler). */
	getAllSlideAnnotations(): SlideAnnotationMap {
		return presentationInkBySlide(this.ink.getSnapshot());
	}

	/** Show the popup toolbar and hide it again after the shared delay (3 s). */
	showToolbarTemporarily(): void {
		this.ink.revealToolbar();
	}

	/** Leave the show: disarm, hide laser and toolbar, keep the strokes for the prompt. */
	resetForExit(): void {
		this.ink.reset();
	}
}
