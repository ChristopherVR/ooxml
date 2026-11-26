import type { PptxSlide } from 'ooxml-core/pptx';
import { collectSlideInk, lassoSlideInk } from 'ooxml-core/pptx/editor/ink';
import type { InkCoordinate, SlideInkStroke } from 'ooxml-core/pptx/editor/ink';

import { DRAW_STAGE } from '../render/draw-canvas';
import { setDrawPencil } from '../render/ink-drawing';
import { isInsideModal } from '../render/slide-keyboard-focus';
import {
	captureStagePoint,
	createLassoGesturePath,
	drawOverlay,
	paintLassoSelection,
	setInkOverlayPoints,
} from './draw-surface';

export function createCanvasInkTools(
	root: HTMLElement,
	handlers: { onSelection: () => void; onDelete: () => void },
) {
	const win = root.ownerDocument.defaultView;
	const listeners = new AbortController();
	const capture = { capture: true, signal: listeners.signal };
	let stage: HTMLElement | undefined;
	let slide: PptxSlide | undefined;
	let strokes: SlideInkStroke[] | undefined;
	let hasInk = false;
	let selection: SlideInkStroke[] = [];
	let lasso: SVGSVGElement | undefined;
	let armed = false;
	let pencil = false;
	let gesture:
		| {
				id: number;
				points: InkCoordinate[];
				project: ReturnType<typeof captureStagePoint>;
				path: SVGPathElement;
		  }
		| undefined;
	let frame: number | undefined;
	const all = () => (strokes ??= slide ? collectSlideInk(slide.elements) : []);
	const clearGesture = () => {
		if (frame !== undefined) win?.cancelAnimationFrame(frame);
		frame = undefined;
		gesture?.path.remove();
		gesture = undefined;
	};
	const showLasso = (target: HTMLElement) => {
		if (!lasso?.isConnected) {
			lasso = drawOverlay(target);
			lasso.dataset.pptxDrawLasso = '';
		}
		return lasso;
	};
	const paintSelection = () => {
		if (!stage || !selection.length) {
			lasso?.remove();
			lasso = undefined;
			return;
		}
		paintLassoSelection(showLasso(stage), selection);
	};
	const cancel = () => {
		armed = false;
		clearGesture();
		selection = [];
		paintSelection();
	};
	const onDown = (event: PointerEvent) => {
		if (
			!armed ||
			!stage ||
			gesture ||
			event.button !== 0 ||
			!(event.target instanceof Node) ||
			!stage.contains(event.target) ||
			(event.target instanceof Element && event.target.closest('[data-pptx-draw-ruler]'))
		) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		selection = [];
		const project = captureStagePoint(stage);
		const points = [project(event)];
		paintSelection();
		gesture = {
			id: event.pointerId,
			project,
			points,
			path: createLassoGesturePath(showLasso(stage), points),
		};
	};
	const onMove = (event: PointerEvent) => {
		if (!stage || gesture?.id !== event.pointerId) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		const point = gesture.project(event);
		const last = gesture.points.at(-1);
		if (last && Math.hypot(last.x - point.x, last.y - point.y) >= 2) {
			if (gesture.points.length === 255) {
				gesture.points = gesture.points.filter((_, index) => index % 2 === 0);
			}
			gesture.points.push(point);
			frame ??= win?.requestAnimationFrame(() => {
				frame = undefined;
				if (gesture) setInkOverlayPoints(gesture.path, gesture.points);
			});
		}
	};
	const onUp = (event: PointerEvent) => {
		if (!stage || gesture?.id !== event.pointerId) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		selection = lassoSlideInk(all(), [...gesture.points, gesture.project(event)]);
		clearGesture();
		paintSelection();
		handlers.onSelection();
	};
	root.addEventListener('pointerdown', onDown, capture);
	win?.addEventListener('pointermove', onMove, capture);
	win?.addEventListener('pointerup', onUp, capture);
	win?.addEventListener(
		'pointercancel',
		(event) => {
			if (gesture?.id === event.pointerId) {
				clearGesture();
				paintSelection();
			}
		},
		capture,
	);
	root.addEventListener(
		'keydown',
		(event) => {
			if (isInsideModal(event)) return;
			if (event.key === 'Escape' && (armed || selection.length)) {
				event.preventDefault();
				event.stopImmediatePropagation();
				cancel();
				handlers.onSelection();
			} else if (
				event.key === 'Delete' &&
				selection.length &&
				(!(event.target instanceof Element) ||
					!event.target.closest('input,select,textarea,[contenteditable=true],dialog'))
			) {
				event.preventDefault();
				event.stopImmediatePropagation();
				handlers.onDelete();
			}
		},
		capture,
	);
	return {
		get stage() {
			return stage;
		},
		get selection() {
			return selection;
		},
		get armed() {
			return armed;
		},
		get hasInk() {
			return hasInk;
		},
		all,
		sync(next: PptxSlide | undefined) {
			const nextStage = next
				? (root.querySelector<HTMLElement>(DRAW_STAGE) ?? undefined)
				: undefined;
			const changed = nextStage !== stage || next?.elements !== slide?.elements;
			if (changed) {
				const sameSlide = next?.id === slide?.id;
				strokes = undefined;
				hasInk =
					next?.elements.some(
						(element) =>
							(element.type === 'ink' && element.inkPaths.length > 0) ||
							(element.type === 'contentPart' && Boolean(element.inkStrokes?.length)),
					) ?? false;
				selection = [];
				clearGesture();
				armed &&= sameSlide;
				if (stage !== nextStage) {
					lasso?.remove();
					lasso = undefined;
					if (stage) setDrawPencil(stage, false);
					stage = nextStage;
					if (stage) setDrawPencil(stage, pencil);
				}
				paintSelection();
			}
			slide = next;
			return changed;
		},
		arm(value: boolean) {
			cancel();
			armed = value;
		},
		pencil(active: boolean) {
			pencil = active;
			if (stage) setDrawPencil(stage, active);
		},
		dispose() {
			listeners.abort();
			cancel();
			if (stage) setDrawPencil(stage, false);
		},
	};
}
