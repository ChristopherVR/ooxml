import type { ViewerController } from './controller';
import {
	MAX_ZOOM,
	MIN_ZOOM,
	pinchDistance,
	pinchMidpoint,
	pinchScrollDelta,
	pointFraction,
	pinchZoom,
	type TouchPoint,
} from './viewer-touch-math';

interface Pinch {
	ids: [number, number];
	startDistance: number;
	startZoom: number;
	/** Where the pinch midpoint fell on the paper (0..1), kept under the fingers while zooming. */
	fraction: TouchPoint;
}
/**
 * Two-finger pinch zoom around the midpoint with simultaneous two-finger pan. It runs before the
 * one-pointer shape and marquee gestures and, once a second finger lands, takes the pointers over.
 */
export class ViewerTouchGestures {
	readonly #touches = new Map<number, TouchPoint>();
	/** Pointers consumed by a pinch stay ignored until they lift, so no finger starts a pan after it. */
	readonly #spent = new Set<number>();
	#pinch: Pinch | undefined;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		private readonly options: { active(): boolean; cancelSingle(): void },
	) {}
	get pinching(): boolean {
		return !!this.#pinch;
	}
	wire(): () => void {
		const view = this.viewport.ownerDocument.defaultView;
		const Abort = view?.AbortController ?? AbortController;
		const events = new Abort();
		const options = { signal: events.signal, capture: true };
		const viewport = this.viewport;
		viewport.addEventListener('pointerdown', (event) => this.#down(event), options);
		viewport.addEventListener('pointermove', (event) => this.#move(event), options);
		for (const type of ['pointerup', 'pointercancel'] as const)
			viewport.addEventListener(type, (event) => this.#up(event), options);
		return () => {
			events.abort();
			this.#touches.clear();
			this.#spent.clear();
			this.#pinch = undefined;
		};
	}
	#down(event: PointerEvent): void {
		if (event.pointerType !== 'touch') return;
		this.#touches.set(event.pointerId, { x: event.clientX, y: event.clientY });
		if (this.#spent.size || this.#pinch) {
			this.#spent.add(event.pointerId);
			event.stopImmediatePropagation();
			return;
		}
		if (this.#touches.size < 2 || !this.options.active()) return;
		const [first, second] = [...this.#touches.keys()] as [number, number];
		const a = this.#touches.get(first)!,
			b = this.#touches.get(second)!;
		const distance = pinchDistance(a, b);
		if (!(distance > 0)) return;
		this.options.cancelSingle();
		const paper = this.viewport.querySelector('svg.paper') ?? this.viewport;
		this.#pinch = {
			ids: [first, second],
			startDistance: distance,
			startZoom: this.controller.state.zoom,
			fraction: pointFraction(pinchMidpoint(a, b), paper.getBoundingClientRect()),
		};
		this.#spent.add(first).add(second);
		for (const id of this.#pinch.ids) this.viewport.setPointerCapture?.(id);
		event.preventDefault();
		event.stopImmediatePropagation();
	}
	#move(event: PointerEvent): void {
		if (event.pointerType !== 'touch' || !this.#touches.has(event.pointerId)) return;
		this.#touches.set(event.pointerId, { x: event.clientX, y: event.clientY });
		if (this.#spent.has(event.pointerId)) event.stopImmediatePropagation();
		const pinch = this.#pinch;
		if (!pinch || !pinch.ids.includes(event.pointerId)) return;
		const a = this.#touches.get(pinch.ids[0])!,
			b = this.#touches.get(pinch.ids[1])!;
		const next = pinchZoom(
			pinch.startZoom,
			pinch.startDistance,
			pinchDistance(a, b),
			MIN_ZOOM,
			MAX_ZOOM,
		);
		this.controller.setZoom(next);
		// Rendering is synchronous: measure the resized paper and scroll its pinch point back under
		// the fingers, which also follows the two-finger pan.
		const paper = this.viewport.querySelector('svg.paper');
		if (paper) {
			const delta = pinchScrollDelta(
				pinch.fraction,
				paper.getBoundingClientRect(),
				pinchMidpoint(a, b),
			);
			this.viewport.scrollLeft += delta.x;
			this.viewport.scrollTop += delta.y;
		}
		event.preventDefault();
	}
	#up(event: PointerEvent): void {
		if (event.pointerType !== 'touch') return;
		this.#touches.delete(event.pointerId);
		if (this.#pinch?.ids.includes(event.pointerId)) this.#pinch = undefined;
		if (this.#spent.delete(event.pointerId)) event.stopImmediatePropagation();
	}
}
