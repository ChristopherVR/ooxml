/**
 * Slide-show ink (the pen, highlighter, eraser and laser pointer of the show toolbar) as a
 * headless controller: the armed tool, the two colours, the strokes drawn on each slide, the
 * stroke in progress, the laser dot, Ctrl+M's "show ink markup" and the popup toolbar's
 * auto-hide.
 *
 * Coordinates are the binding's own (slide pixels in most, unit square in vanilla); the
 * controller never converts them, only the eraser radius is in the same units
 * ({@link PresentationInkControllerOptions.eraserRadius}).
 *
 * Choices where the bindings had drifted (see `docs/structure-plan.md` section 5):
 *
 *  - choosing the armed tool again disarms it ({@link PresentationInkController.toggleTool});
 *    {@link PresentationInkController.setTool} arms without toggling;
 *  - pen and highlighter keep separate colours; the highlighter is 14 wide at 0.4 opacity, the
 *    pen 2.5 wide and opaque;
 *  - a stroke needs two points; the eraser drops every stroke on the slide with a point within
 *    the radius;
 *  - "Erase all ink on slide" clears the current slide only, {@link PresentationInkController.clearAll}
 *    every slide;
 *  - leaving the show ({@link PresentationInkController.reset}) disarms the tool and hides the
 *    laser and toolbar but keeps the strokes for the keep-or-discard prompt.
 *
 * @module render/controller/presentation-ink-controller
 */

import type { PptxSlide } from 'ooxml-core/pptx';

import { strokeToInkElement } from '../ink-drawing';
import type {
	PresentationInkPoint,
	PresentationInkStroke,
	PresentationPointerTool,
} from '../presentation-session-types';
import { createControllerStore } from './controller';
import type { PptxController } from './controller';

/** One slide-show stroke: the shared session stroke plus its drawing opacity. */
export interface PresentationAnnotationStroke extends PresentationInkStroke {
	readonly opacity: number;
}

export interface PresentationInkState {
	readonly tool: PresentationPointerTool;
	readonly penColor: string;
	readonly highlighterColor: string;
	/** The slide strokes are drawn on and erased from. */
	readonly slideIndex: number;
	/** Finished strokes on every slide, in drawing order. */
	readonly strokes: readonly PresentationAnnotationStroke[];
	/** The stroke being drawn, or null. */
	readonly current: PresentationAnnotationStroke | null;
	/** The laser dot, or null when hidden. */
	readonly laser: PresentationInkPoint | null;
	/** Ctrl+M: whether ink is drawn. Hiding keeps the strokes. */
	readonly markupVisible: boolean;
	/** The popup show toolbar is revealed. */
	readonly toolbarVisible: boolean;
}

/** Timeout seam; defaults to the global `setTimeout` / `clearTimeout`. */
export interface PresentationInkTimers {
	setTimeout(handler: () => void, ms: number): unknown;
	clearTimeout(id: unknown): void;
}

export interface PresentationInkControllerOptions {
	/** Eraser hit radius in stroke coordinates; defaults to 16. */
	eraserRadius?: number;
	/** How long the popup toolbar stays after the last reveal; defaults to 3000 ms. */
	toolbarHideMs?: number;
	timers?: PresentationInkTimers;
	/** Stroke id factory; defaults to a per-controller sequence. */
	createId?(): string;
}

export interface PresentationInkController extends PptxController<PresentationInkState> {
	/** Arm `tool`, or disarm it when it is already armed. */
	toggleTool(tool: PresentationPointerTool): void;
	/** Arm `tool` (no toggle). */
	setTool(tool: PresentationPointerTool): void;
	setPenColor(color: string): void;
	setHighlighterColor(color: string): void;
	/** The show moved to `slideIndex`; a stroke in progress is dropped. */
	setSlide(slideIndex: number): void;
	/** Pointer down: starts a stroke (pen, highlighter), erases (eraser) or moves the laser. */
	pointerDown(point: PresentationInkPoint): void;
	/** Pointer move: extends the stroke, keeps erasing or moves the laser. */
	pointerMove(point: PresentationInkPoint): void;
	/** Pointer up or cancel: commits the stroke and ends an erase. */
	pointerUp(): void;
	/** The pointer left the stage: hides the laser. */
	pointerLeave(): void;
	/** Erase at `point` regardless of a gesture (eraser armed only). */
	eraseAt(point: PresentationInkPoint): void;
	/** Clear the current slide's strokes. */
	clearSlide(): void;
	/** Clear every slide's strokes. */
	clearAll(): void;
	/** Replace every stroke (a presenter snapshot arriving, a restored session). */
	setStrokes(strokes: readonly PresentationInkStroke[]): void;
	setMarkupVisible(visible: boolean): void;
	/** Show the popup toolbar and hide it after the configured delay. */
	revealToolbar(): void;
	setToolbarVisible(visible: boolean): void;
	/** Leave the show: disarm, hide laser and toolbar, keep the strokes. */
	reset(): void;
	dispose(): void;
}

export const PRESENTATION_PEN_WIDTH = 2.5;
export const PRESENTATION_HIGHLIGHTER_WIDTH = 14;
export const PRESENTATION_HIGHLIGHTER_OPACITY = 0.4;
export const PRESENTATION_ERASER_RADIUS = 16;
/** The defaults of the show toolbar's palettes (`PEN_COLORS[0]`, `HIGHLIGHTER_COLORS[0]`). */
export const PRESENTATION_PEN_COLOR = '#ff0000';
export const PRESENTATION_HIGHLIGHTER_COLOR = '#ffff00';

const defaultTimers: PresentationInkTimers = {
	setTimeout: (handler, ms) => setTimeout(handler, ms),
	clearTimeout: (id) => clearTimeout(id as ReturnType<typeof setTimeout>),
};

/** The strokes drawn on `slideIndex`. */
export function presentationInkOnSlide(
	state: PresentationInkState,
	slideIndex: number = state.slideIndex,
): PresentationAnnotationStroke[] {
	return state.strokes.filter((stroke) => stroke.slideIndex === slideIndex);
}

/** The strokes grouped by slide (for the keep-or-discard prompt). */
export function presentationInkBySlide(
	state: PresentationInkState,
): Map<number, PresentationAnnotationStroke[]> {
	const map = new Map<number, PresentationAnnotationStroke[]>();
	for (const stroke of state.strokes) {
		map.set(stroke.slideIndex, [...(map.get(stroke.slideIndex) ?? []), stroke]);
	}
	return map;
}

/** How many slides carry ink. */
export function presentationInkSlideCount(state: PresentationInkState): number {
	return new Set(state.strokes.map((stroke) => stroke.slideIndex)).size;
}

/**
 * Append the strokes to their slides as ink elements ("Keep" in the exit prompt). `scale`
 * maps stroke coordinates to slide pixels (vanilla's unit-square strokes pass the canvas
 * size). Returns `slides` itself when nothing was added.
 */
export function applyPresentationInk(
	slides: readonly PptxSlide[],
	strokes: readonly PresentationInkStroke[],
	scale: { width: number; height: number } = { width: 1, height: 1 },
): readonly PptxSlide[] {
	let changed = false;
	const next = slides.map((slide, slideIndex) => {
		const additions = strokes
			.filter((stroke) => stroke.slideIndex === slideIndex && stroke.points.length > 1)
			.map((stroke) =>
				strokeToInkElement({
					points: stroke.points.map((point) => ({
						x: point.x * scale.width,
						y: point.y * scale.height,
					})),
					color: stroke.color,
					width: stroke.width,
					tool: stroke.tool,
				}),
			)
			.filter((ink) => ink !== null);
		if (additions.length === 0) {
			return slide;
		}
		changed = true;
		return { ...slide, elements: [...slide.elements, ...additions] };
	});
	return changed ? next : slides;
}

const INITIAL: PresentationInkState = {
	tool: 'none',
	penColor: PRESENTATION_PEN_COLOR,
	highlighterColor: PRESENTATION_HIGHLIGHTER_COLOR,
	slideIndex: 0,
	strokes: [],
	current: null,
	laser: null,
	markupVisible: true,
	toolbarVisible: false,
};

/** Create the slide-show ink controller. */
export function createPresentationInkController(
	options: PresentationInkControllerOptions = {},
): PresentationInkController {
	const radius = options.eraserRadius ?? PRESENTATION_ERASER_RADIUS;
	const hideMs = options.toolbarHideMs ?? 3000;
	const timers = options.timers ?? defaultTimers;
	let sequence = 0;
	const createId = options.createId ?? (() => `presentation-stroke-${++sequence}`);
	const store = createControllerStore<PresentationInkState>(INITIAL);
	let erasing = false;
	let toolbarTimer: unknown = null;

	const patch = (next: Partial<PresentationInkState>): void =>
		store.update((state) => ({ ...state, ...next }));
	const clearToolbarTimer = (): void => {
		if (toolbarTimer !== null) {
			timers.clearTimeout(toolbarTimer);
			toolbarTimer = null;
		}
	};
	const arm = (tool: PresentationPointerTool): void => {
		erasing = false;
		store.update((state) =>
			state.tool === tool
				? state
				: {
						...state,
						tool,
						current: tool === 'pen' || tool === 'highlighter' ? state.current : null,
						laser: tool === 'laser' ? state.laser : null,
					},
		);
	};
	const erase = (point: PresentationInkPoint): void =>
		store.update((state) => {
			const r2 = radius * radius;
			const strokes = state.strokes.filter(
				(stroke) =>
					stroke.slideIndex !== state.slideIndex ||
					!stroke.points.some((p) => (p.x - point.x) ** 2 + (p.y - point.y) ** 2 < r2),
			);
			return strokes.length === state.strokes.length ? state : { ...state, strokes };
		});

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		toggleTool(tool) {
			arm(store.getSnapshot().tool === tool ? 'none' : tool);
		},
		setTool: arm,
		setPenColor: (penColor) => patch({ penColor }),
		setHighlighterColor: (highlighterColor) => patch({ highlighterColor }),
		setSlide(slideIndex) {
			erasing = false;
			store.update((state) =>
				state.slideIndex === slideIndex ? state : { ...state, slideIndex, current: null },
			);
		},
		pointerDown(point) {
			const state = store.getSnapshot();
			if (state.tool === 'eraser') {
				erasing = true;
				erase(point);
			} else if (state.tool === 'laser') {
				patch({ laser: point });
			} else if (state.tool === 'pen' || state.tool === 'highlighter') {
				const pen = state.tool === 'pen';
				patch({
					current: {
						id: createId(),
						slideIndex: state.slideIndex,
						tool: state.tool,
						color: pen ? state.penColor : state.highlighterColor,
						width: pen ? PRESENTATION_PEN_WIDTH : PRESENTATION_HIGHLIGHTER_WIDTH,
						opacity: pen ? 1 : PRESENTATION_HIGHLIGHTER_OPACITY,
						points: [point],
					},
				});
			}
		},
		pointerMove(point) {
			const state = store.getSnapshot();
			if (state.tool === 'laser') {
				patch({ laser: point });
			} else if (state.tool === 'eraser') {
				if (erasing) {
					erase(point);
				}
			} else if (state.current) {
				patch({ current: { ...state.current, points: [...state.current.points, point] } });
			}
		},
		pointerUp() {
			erasing = false;
			store.update((state) => {
				const stroke = state.current;
				if (!stroke) {
					return state;
				}
				return stroke.points.length < 2
					? { ...state, current: null }
					: { ...state, current: null, strokes: [...state.strokes, stroke] };
			});
		},
		pointerLeave() {
			erasing = false;
			patch({ laser: null });
		},
		eraseAt(point) {
			if (store.getSnapshot().tool === 'eraser') {
				erase(point);
			}
		},
		clearSlide() {
			erasing = false;
			store.update((state) => ({
				...state,
				current: null,
				strokes: state.strokes.filter((stroke) => stroke.slideIndex !== state.slideIndex),
			}));
		},
		clearAll() {
			erasing = false;
			patch({ current: null, strokes: [] });
		},
		setStrokes(strokes) {
			patch({
				strokes: strokes.map((stroke) => ({
					...stroke,
					opacity:
						'opacity' in stroke && typeof stroke.opacity === 'number'
							? stroke.opacity
							: stroke.tool === 'highlighter'
								? PRESENTATION_HIGHLIGHTER_OPACITY
								: 1,
				})),
			});
		},
		setMarkupVisible: (markupVisible) => patch({ markupVisible }),
		revealToolbar() {
			patch({ toolbarVisible: true });
			clearToolbarTimer();
			toolbarTimer = timers.setTimeout(() => {
				toolbarTimer = null;
				patch({ toolbarVisible: false });
			}, hideMs);
		},
		setToolbarVisible(toolbarVisible) {
			clearToolbarTimer();
			patch({ toolbarVisible });
		},
		reset() {
			erasing = false;
			clearToolbarTimer();
			patch({ tool: 'none', current: null, laser: null, toolbarVisible: false });
		},
		dispose: clearToolbarTimer,
	};
}
