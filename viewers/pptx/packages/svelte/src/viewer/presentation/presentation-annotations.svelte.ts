import {
	applyPresentationInk,
	createPresentationInkController,
	presentationInkBySlide,
	presentationInkOnSlide,
	presentationInkSlideCount,
} from 'ooxml-ui/pptx';
import type {
	InkPoint,
	PresentationAnnotationStroke,
	PresentationInkController,
	PresentationPointerTool,
} from 'ooxml-ui/pptx';

import type { EditorState } from '../editor/editor-state.svelte';
import { controllerState } from '../state/controller-state';

export type PresentationInkTool = PresentationPointerTool;
export type PresentationStroke = PresentationAnnotationStroke;

/**
 * Slide-show ink for the Svelte binding: the shared
 * `createPresentationInkController` read through the store adapter. The
 * properties keep the names the templates bind to; every change goes through
 * the controller, which owns the tool, colours and per-slide strokes.
 */
export class PresentationAnnotations {
	readonly controller: PresentationInkController = createPresentationInkController();
	readonly #state = controllerState(this.controller);

	get tool(): PresentationInkTool {
		return this.#state.current.tool;
	}
	set tool(tool: PresentationInkTool) {
		this.controller.setTool(tool);
	}
	get penColor(): string {
		return this.#state.current.penColor;
	}
	set penColor(color: string) {
		this.controller.setPenColor(color);
	}
	get highlighterColor(): string {
		return this.#state.current.highlighterColor;
	}
	set highlighterColor(color: string) {
		this.controller.setHighlighterColor(color);
	}
	get current(): PresentationStroke | null {
		return this.#state.current.current;
	}
	get laser(): InkPoint | null {
		return this.#state.current.laser;
	}
	set laser(point: InkPoint | null) {
		if (point === null) {
			this.controller.pointerLeave();
		}
	}
	get bySlide(): Map<number, PresentationStroke[]> {
		return presentationInkBySlide(this.#state.current);
	}
	get count(): number {
		return this.#state.current.strokes.length;
	}
	get slideCount(): number {
		return presentationInkSlideCount(this.#state.current);
	}
	strokes(index: number): readonly PresentationStroke[] {
		return presentationInkOnSlide(this.#state.current, index);
	}

	pointerDown(index: number, point: InkPoint): void {
		this.controller.setSlide(index);
		this.controller.pointerDown(point);
	}
	pointerMove(index: number, point: InkPoint): void {
		this.controller.setSlide(index);
		this.controller.pointerMove(point);
	}
	pointerUp(index: number): void {
		this.controller.setSlide(index);
		this.controller.pointerUp();
	}
	/** "Erase All Ink on Slide": the slide being shown. */
	clearSlide(): void {
		this.controller.clearSlide();
	}
	/** Drop every stroke and disarm (discarding at the end of the show). */
	clear(): void {
		this.controller.clearAll();
		this.controller.reset();
	}
	keep(editor: EditorState): void {
		const slides = applyPresentationInk(editor.slides, this.controller.getSnapshot().strokes);
		if (slides !== editor.slides) {
			editor.commitSlides([...slides]);
		}
		this.clear();
	}
}
