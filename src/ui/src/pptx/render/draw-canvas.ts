import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';

import type { ViewerMode } from '../types';

export const DRAW_CANVAS_SCOPE = '[data-pptx-theme-root]';
export const DRAW_CANVAS_CHANGE = 'pptx-draw-canvas-change';
export const DRAW_STAGE = '[data-pptx-viewport] [aria-roledescription="slide"]';

interface DrawCanvasState {
	slides: readonly PptxSlide[];
	index: number;
	mode: ViewerMode;
	canEdit: boolean;
	editingTemplate?: boolean;
}

function canEditDraw({ mode, canEdit, editingTemplate }: DrawCanvasState): boolean {
	return canEdit && mode === 'edit' && !editingTemplate;
}

export interface DrawCanvasPort {
	readonly missingCommit?: boolean;
	getSlide(): PptxSlide | undefined;
	replaceElements(slide: PptxSlide, elements: PptxElement[], selectedId: string | null): void;
}

export function createDrawCanvasPort(
	read: () => DrawCanvasState,
	getCommit?: () => ((slides: PptxSlide[], selectedId: string | null) => void) | undefined,
): DrawCanvasPort {
	return {
		get missingCommit() {
			return !getCommit?.();
		},
		getSlide() {
			const state = read();
			return canEditDraw(state) && getCommit?.() ? state.slides[state.index] : undefined;
		},
		replaceElements(slide, elements, selectedId) {
			const state = read();
			const commit = getCommit?.();
			if (!commit || !canEditDraw(state)) return;
			const { slides, index } = state;
			const current = slides[index];
			if (current?.id !== slide.id || current.elements !== slide.elements) return;
			const next = [...slides];
			next[index] = { ...current, elements };
			commit(next, selectedId);
		},
	};
}

export interface DrawPenPreset {
	kind: 'pencil' | 'pen' | 'highlighter' | 'actionPen';
	color: string;
	width: number;
}

export interface PptxDrawHostOptions {
	penPreferences?: {
		read(): unknown;
		write(presets: readonly DrawPenPreset[]): void;
	};
	/** Called only on an explicit Ink to Math request, with selected slide-local coordinates. */
	recognizeMath?: (
		strokes: ReadonlyArray<ReadonlyArray<Readonly<{ x: number; y: number }>>>,
		signal: AbortSignal,
	) => Promise<string | undefined>;
}

const canvases = new WeakMap<HTMLElement, DrawCanvasPort>();
const options = new WeakMap<HTMLElement, PptxDrawHostOptions>();

export function notifyDrawCanvas(root: HTMLElement): void {
	(root.closest<HTMLElement>(DRAW_CANVAS_SCOPE) ?? root).dispatchEvent(
		new Event(DRAW_CANVAS_CHANGE),
	);
}

export function attachDrawCanvas(root: HTMLElement, port: DrawCanvasPort): () => void {
	root = root.closest<HTMLElement>(DRAW_CANVAS_SCOPE) ?? root;
	canvases.set(root, port);
	notifyDrawCanvas(root);
	return () => {
		if (canvases.get(root) === port) {
			canvases.delete(root);
			notifyDrawCanvas(root);
		}
	};
}

export function configurePptxDraw(root: HTMLElement, settings: PptxDrawHostOptions): () => void {
	root = root.closest<HTMLElement>(DRAW_CANVAS_SCOPE) ?? root;
	options.set(root, settings);
	notifyDrawCanvas(root);
	return () => {
		if (options.get(root) === settings) {
			options.delete(root);
			notifyDrawCanvas(root);
		}
	};
}

export function drawCanvasPort(root: HTMLElement) {
	return canvases.get(root);
}

export function drawHostOptions(root: HTMLElement) {
	return options.get(root);
}
