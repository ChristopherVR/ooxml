import type { PptxSlide } from 'ooxml-core/pptx';
import { collectSlideInk, lassoSlideInk, projectInkPoint } from 'ooxml-core/pptx/editor/ink';
import type { InkCoordinate, SlideInkStroke } from 'ooxml-core/pptx/editor/ink';

import { DRAW_STAGE } from '../render/draw-canvas';
import { setDrawPencil, setDrawRulerProjection } from '../render/ink-drawing';
import { drawOverlay, inkOverlayPath, paintLasso, stagePoint } from './draw-surface';

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
	let ruler: SVGSVGElement | undefined;
	let armed = false;
	let pencil = false;
	let rulerVisible = false;
	let angle = 0;
	let gesture: { id: number; points: InkCoordinate[] } | undefined;
	const all = () => (strokes ??= slide ? collectSlideInk(slide.elements) : []);
	const removeRuler = () => {
		if (stage) {
			setDrawRulerProjection(stage);
		}
		ruler?.remove();
		ruler = undefined;
	};
	const paintRuler = () => {
		removeRuler();
		if (!stage || !rulerVisible) {
			return;
		}
		const x = stage.offsetWidth / 2;
		const y = stage.offsetHeight / 2;
		const length = Math.min(stage.offsetWidth, stage.offsetHeight) * 0.43;
		const radians = (angle * Math.PI) / 180;
		const dx = length * Math.cos(radians);
		const dy = length * Math.sin(radians);
		const a = { x: x - dx, y: y - dy };
		const b = { x: x + dx, y: y + dy };
		ruler = drawOverlay(stage);
		ruler.dataset.pptxDrawRuler = '';
		const edge = inkOverlayPath(ruler, [a, b]);
		edge.setAttribute('stroke', '#416685');
		edge.setAttribute('stroke-width', '3');
		const ticks = inkOverlayPath(ruler, [a, b]);
		ticks.setAttribute('stroke', '#d5e8fb');
		ticks.setAttribute('stroke-width', '14');
		ticks.setAttribute('stroke-dasharray', '2 10');
		ticks.setAttribute(
			'transform',
			`translate(${-Math.sin(radians) * 8} ${Math.cos(radians) * 8})`,
		);
		setDrawRulerProjection(stage, (point) => {
			const projection = projectInkPoint(point, a, b);
			return projection.offset >= 0 &&
				projection.offset <= 1 &&
				Math.hypot(point.x - projection.point.x, point.y - projection.point.y) <= 16
				? projection.point
				: point;
		});
	};
	const paint = () => {
		if (!stage || (!gesture && !selection.length)) {
			lasso?.remove();
			lasso = undefined;
			return;
		}
		if (!lasso?.isConnected) {
			lasso = drawOverlay(stage);
			lasso.dataset.pptxDrawLasso = '';
		}
		paintLasso(lasso, gesture?.points ?? [], selection);
	};
	const cancel = () => {
		armed = false;
		gesture = undefined;
		selection = [];
		paint();
	};
	const onDown = (event: PointerEvent) => {
		if (
			!armed ||
			!stage ||
			gesture ||
			event.button !== 0 ||
			!(event.target instanceof Node) ||
			!stage.contains(event.target)
		) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		selection = [];
		gesture = { id: event.pointerId, points: [stagePoint(stage, event)] };
		paint();
	};
	const onMove = (event: PointerEvent) => {
		if (!stage || gesture?.id !== event.pointerId) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		const point = stagePoint(stage, event);
		const last = gesture.points.at(-1);
		if (last && Math.hypot(last.x - point.x, last.y - point.y) >= 2) {
			if (gesture.points.length === 255) {
				gesture.points = gesture.points.filter((_, index) => index % 2 === 0);
			}
			gesture.points.push(point);
			paint();
		}
	};
	const onUp = (event: PointerEvent) => {
		if (!stage || gesture?.id !== event.pointerId) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		selection = lassoSlideInk(all(), [...gesture.points, stagePoint(stage, event)]);
		gesture = undefined;
		paint();
		handlers.onSelection();
	};
	root.addEventListener('pointerdown', onDown, capture);
	win?.addEventListener('pointermove', onMove, capture);
	win?.addEventListener('pointerup', onUp, capture);
	win?.addEventListener(
		'pointercancel',
		(event) => {
			if (gesture?.id === event.pointerId) {
				gesture = undefined;
				paint();
			}
		},
		capture,
	);
	root.addEventListener(
		'keydown',
		(event) => {
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
		get rulerVisible() {
			return rulerVisible;
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
				gesture = undefined;
				armed &&= sameSlide;
				if (stage !== nextStage) {
					if (stage) setDrawPencil(stage, false);
					removeRuler();
					stage = nextStage;
					if (stage) setDrawPencil(stage, pencil);
					paintRuler();
				}
				paint();
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
		ruler(visible: boolean, degrees: number) {
			rulerVisible = visible;
			angle = degrees;
			paintRuler();
		},
		dispose() {
			listeners.abort();
			cancel();
			if (stage) setDrawPencil(stage, false);
			removeRuler();
		},
	};
}
