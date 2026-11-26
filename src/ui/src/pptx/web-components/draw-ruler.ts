import { projectInkPoint, type InkCoordinate } from 'ooxml-core/pptx/editor/ink';

import { DRAW_CANVAS_CHANGE, DRAW_STAGE, drawCanvasPort } from '../render/draw-canvas';
import { setDrawRulerProjection } from '../render/ink-drawing';
import {
	captureStagePoint,
	drawOverlay,
	inkOverlayPath,
	setInkOverlayPoints,
} from './draw-surface';

const rulers = new WeakMap<HTMLElement, ReturnType<typeof createDrawRuler>>();

export function drawRulerFor(root: HTMLElement) {
	let ruler = rulers.get(root);
	if (!ruler) {
		ruler = createDrawRuler(root);
		rulers.set(root, ruler);
	}
	return ruler;
}

/** The ruler belongs to the viewer canvas and can outlive the Draw tab. */
function createDrawRuler(root: HTMLElement) {
	let visible = false;
	let angle = 0;
	let center: InkCoordinate | undefined;
	let labels = { name: '', instructions: '' };
	let surface: ReturnType<typeof mountRuler> | undefined;
	let listeners: AbortController | undefined;
	let observer: MutationObserver | undefined;
	const changes = new Set<() => void>();
	const emit = () => changes.forEach((listener) => listener());

	function change(position: InkCoordinate, degrees: number) {
		const stage = surface?.stage;
		if (!stage) return;
		center = {
			x: Math.max(0, Math.min(stage.offsetWidth, position.x)),
			y: Math.max(0, Math.min(stage.offsetHeight, position.y)),
		};
		angle = degrees > 180 ? degrees - 360 : degrees < -180 ? degrees + 360 : degrees;
		surface?.paint();
		emit();
	}
	function syncCanvas() {
		const port = drawCanvasPort(root);
		if (!port) {
			stopWatching();
			visible = false;
			rulers.delete(root);
		}
		const stage = visible && port?.getSlide() ? root.querySelector<HTMLElement>(DRAW_STAGE) : null;
		if (stage === surface?.stage && surface.svg.isConnected) return;
		surface?.dispose();
		surface = stage
			? mountRuler(stage, {
					center: () => center ?? { x: stage.offsetWidth / 2, y: stage.offsetHeight / 2 },
					angle: () => angle,
					labels: () => labels,
					change,
				})
			: undefined;
	}
	function stopWatching() {
		listeners?.abort();
		listeners = undefined;
		observer?.disconnect();
		observer = undefined;
	}
	return {
		get visible() {
			return visible;
		},
		get angle() {
			return angle;
		},
		setLabels(name: string, instructions: string) {
			labels = { name, instructions };
			surface?.label();
		},
		set(nextVisible: boolean, degrees: number) {
			visible = nextVisible;
			angle = degrees;
			if (visible && !listeners) {
				listeners = new AbortController();
				root.addEventListener(DRAW_CANVAS_CHANGE, syncCanvas, { signal: listeners.signal });
				observer = new MutationObserver(syncCanvas);
				observer.observe(root, { childList: true, subtree: true });
			} else if (!visible) {
				stopWatching();
			}
			syncCanvas();
			surface?.paint();
			emit();
		},
		subscribe(listener: () => void) {
			changes.add(listener);
			return () => {
				changes.delete(listener);
			};
		},
	};
}

function mountRuler(
	stage: HTMLElement,
	state: {
		center: () => InkCoordinate;
		angle: () => number;
		labels: () => { name: string; instructions: string };
		change: (position: InkCoordinate, angle: number) => void;
	},
) {
	const svg = drawOverlay(stage);
	svg.dataset.pptxDrawRuler = '';
	svg.removeAttribute('aria-hidden');
	svg.setAttribute('role', 'group');
	svg.setAttribute('tabindex', '0');
	svg.style.outlineColor = 'var(--office-ring, Highlight)';
	svg.style.outlineOffset = '-2px';
	const body = inkOverlayPath(svg, []);
	body.setAttribute('aria-hidden', 'true');
	body.setAttribute('stroke', 'var(--office-surface, Canvas)');
	body.setAttribute('stroke-width', '24');
	body.style.pointerEvents = 'stroke';
	body.style.cursor = 'grab';
	body.style.touchAction = 'none';
	const edge = inkOverlayPath(svg, []);
	edge.setAttribute('aria-hidden', 'true');
	edge.setAttribute('stroke', 'var(--office-foreground, CanvasText)');
	edge.setAttribute('stroke-width', '2');
	const ticks = inkOverlayPath(svg, []);
	ticks.setAttribute('aria-hidden', 'true');
	ticks.setAttribute('stroke', 'var(--office-foreground, CanvasText)');
	ticks.setAttribute('stroke-width', '7');
	ticks.setAttribute('stroke-dasharray', '1 9');

	const label = () => {
		const { name, instructions } = state.labels();
		svg.setAttribute('aria-label', name);
		svg.setAttribute('aria-description', instructions);
	};
	const paint = () => {
		const { x, y } = state.center();
		const length = Math.min(stage.offsetWidth, stage.offsetHeight) * 0.43;
		const radians = (state.angle() * Math.PI) / 180;
		const dx = Math.cos(radians);
		const dy = Math.sin(radians);
		const a = { x: x - length * dx, y: y - length * dy };
		const b = { x: x + length * dx, y: y + length * dy };
		svg.setAttribute('viewBox', `0 0 ${stage.offsetWidth} ${stage.offsetHeight}`);
		for (const path of [edge, ticks, body]) setInkOverlayPoints(path, [a, b]);
		body.setAttribute('transform', `translate(${-dy * 14} ${dx * 14})`);
		ticks.setAttribute('transform', `translate(${-dy * 6} ${dx * 6})`);
		setDrawRulerProjection(stage, (point) => {
			const projection = projectInkPoint(point, a, b);
			return projection.offset >= 0 &&
				projection.offset <= 1 &&
				Math.hypot(point.x - projection.point.x, point.y - projection.point.y) <= 16
				? projection.point
				: point;
		});
	};
	label();
	paint();
	const resize = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(paint) : undefined;
	resize?.observe(stage);
	let drag:
		| { id: number; project: ReturnType<typeof captureStagePoint>; offset: InkCoordinate }
		| undefined;
	body.addEventListener('pointerdown', (event) => {
		if (event.button !== 0 || drag) return;
		event.preventDefault();
		event.stopPropagation();
		svg.focus();
		const project = captureStagePoint(stage);
		const start = project(event);
		const center = state.center();
		drag = {
			id: event.pointerId,
			project,
			offset: { x: center.x - start.x, y: center.y - start.y },
		};
		body.setPointerCapture(event.pointerId);
		body.style.cursor = 'grabbing';
	});
	body.addEventListener('pointermove', (event) => {
		if (drag?.id !== event.pointerId) return;
		const { x, y } = drag.project(event);
		state.change({ x: x + drag.offset.x, y: y + drag.offset.y }, state.angle());
	});
	body.addEventListener('lostpointercapture', () => {
		drag = undefined;
		body.style.cursor = 'grab';
	});
	svg.addEventListener('keydown', (event) => {
		const delta = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[
			event.key
		];
		if (!delta || event.altKey || event.ctrlKey || event.metaKey) return;
		const [dx = 0, dy = 0] = delta;
		event.preventDefault();
		event.stopPropagation();
		const center = state.center();
		if (event.shiftKey) state.change(center, state.angle() + dx + dy);
		else state.change({ x: center.x + 10 * dx, y: center.y + 10 * dy }, state.angle());
	});
	return {
		stage,
		svg,
		paint,
		label,
		dispose() {
			resize?.disconnect();
			setDrawRulerProjection(stage);
			svg.remove();
		},
	};
}
