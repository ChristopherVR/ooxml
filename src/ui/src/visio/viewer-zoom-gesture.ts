import type { ViewerController } from './controller';
import { clampZoom, pinchScrollDelta, type TouchPoint } from './viewer-touch-math';

/**
 * What a zoom keeps in place, as an immediate zoom did: `scroll` keeps the drawing window's
 * scroll offsets (the browser clamps them as the page shrinks), `point` keeps the page position
 * at `fraction` of the paper under the screen `point` (a pinch).
 */
export type ZoomAnchor =
	| { readonly kind: 'scroll' }
	| { readonly kind: 'point'; readonly fraction: TouchPoint; readonly point: TouchPoint };

/** The drawing window measured when a gesture starts, in CSS pixels. */
export interface ZoomLayout {
	/** Screen position of the viewport's padding box. */
	readonly origin: TouchPoint;
	/** The viewport's client size (inside its scroll bars). */
	readonly client: TouchPoint;
	readonly padding: {
		readonly left: number;
		readonly right: number;
		readonly top: number;
		readonly bottom: number;
	};
	/** Paper size at zoom 1. */
	readonly page: TouchPoint;
}

interface Gesture {
	readonly paper: SVGSVGElement;
	/** The zoom the paper is laid out at. */
	readonly base: number;
	readonly layout: ZoomLayout;
	/** Scroll offsets and untransformed paper position when the gesture started. */
	readonly scroll: TouchPoint;
	readonly start: TouchPoint;
	target: number;
	/** The smallest zoom the gesture passed through: an immediate zoom clamped scrolling there. */
	lowest: number;
	anchor: ZoomAnchor;
}

const SCROLL: ZoomAnchor = { kind: 'scroll' };

/** The furthest the viewport scrolls at `zoom` (the paper and the viewport's padding). */
export function scrollRange(layout: ZoomLayout, zoom: number): TouchPoint {
	const { client, padding, page } = layout;
	return {
		x: Math.max(0, page.x * zoom + padding.left + padding.right - client.x),
		y: Math.max(0, page.y * zoom + padding.top + padding.bottom - client.y),
	};
}

/** Where the paper's top-left sits inside the scrolled content: centred until it overflows. */
function paperOffset(layout: ZoomLayout, zoom: number): TouchPoint {
	const { client, padding, page } = layout;
	return {
		x: padding.left + Math.max(0, (client.x - padding.left - padding.right - page.x * zoom) / 2),
		y: padding.top + Math.max(0, (client.y - padding.top - padding.bottom - page.y * zoom) / 2),
	};
}

/** Scroll offsets clamped to the range at `zoom`, or kept when `zoom` is undefined. */
function keptScroll(scroll: TouchPoint, layout: ZoomLayout, zoom: number | undefined): TouchPoint {
	if (zoom === undefined) return scroll;
	const range = scrollRange(layout, zoom);
	return { x: Math.min(scroll.x, range.x), y: Math.min(scroll.y, range.y) };
}

/**
 * Screen position of the paper's top-left once a gesture at `target` has been rendered.
 * `clampAt` is the smallest zoom below the laid-out one that the gesture reached: every immediate
 * step kept the scroll offsets, and the smallest page clamped them the most.
 */
export function settledOrigin(
	layout: ZoomLayout,
	scroll: TouchPoint,
	target: number,
	clampAt: number | undefined,
	anchor: ZoomAnchor,
): TouchPoint {
	if (anchor.kind === 'point')
		return {
			x: anchor.point.x - anchor.fraction.x * layout.page.x * target,
			y: anchor.point.y - anchor.fraction.y * layout.page.y * target,
		};
	const kept = keptScroll(scroll, layout, clampAt);
	const offset = paperOffset(layout, target);
	return { x: layout.origin.x + offset.x - kept.x, y: layout.origin.y + offset.y - kept.y };
}

/**
 * Interactive zoom (Ctrl+wheel, pinch, the zoom slider, the + and - keys). Re-rendering a large
 * page costs a style, layout and paint pass of every shape, so while the gesture runs the paper is
 * only scaled with a CSS transform, once per frame, to where the zoom will put it; when input
 * stops for `settleMs` (or the gesture ends) the controller takes the zoom and the page renders
 * once, so text, line widths and hit testing are exact again. The end zoom and scroll position are
 * those the immediate zoom reached.
 */
export class ViewerZoomGesture {
	/**
	 * Shows the previewed zoom elsewhere (the status bar's zoom slider) and moves what is drawn
	 * outside the paper with it, after each preview frame.
	 */
	onPreview: ((zoom: number) => void) | undefined;
	#gesture: Gesture | undefined;
	#frame = 0;
	#timer: ReturnType<typeof setTimeout> | undefined;
	constructor(
		private readonly viewport: HTMLElement,
		private readonly controller: ViewerController,
		// Long enough for the pause between wheel notches, so one turn of the wheel renders once.
		private readonly settleMs = 300,
	) {}
	/** The zoom the user is looking at: a running gesture's, else the controller's. */
	get zoom(): number {
		return this.#gesture?.target ?? this.controller.state.zoom;
	}
	/** Zoom to `zoom` interactively, keeping `anchor` in place. */
	zoomTo(zoom: number, anchor: ZoomAnchor = SCROLL): void {
		const target = clampZoom(zoom);
		const gesture = this.#gesture ?? this.#start();
		if (!gesture) {
			this.controller.setZoom(target);
			return;
		}
		gesture.target = target;
		gesture.lowest = Math.min(gesture.lowest, target);
		gesture.anchor = anchor;
		const view = this.viewport.ownerDocument.defaultView;
		if (!this.#frame && view?.requestAnimationFrame)
			this.#frame = view.requestAnimationFrame(() => {
				this.#frame = 0;
				this.#paint();
			});
		if (!this.#frame) this.#paint();
		this.#settleLater();
	}
	/** Settle `settleMs` after the last input and the last preview frame, whichever is later. */
	#settleLater(): void {
		clearTimeout(this.#timer);
		this.#timer = setTimeout(() => this.settle(), this.settleMs);
	}
	/** Render the pending zoom now: the gesture ended, or something is about to use the page. */
	settle(): void {
		const gesture = this.#end();
		if (!gesture || this.controller.state.zoom !== gesture.base) return;
		this.controller.setZoom(gesture.target);
		if (!gesture.paper.isConnected) return;
		const { viewport } = this;
		if (gesture.anchor.kind === 'point') {
			const delta = pinchScrollDelta(
				gesture.anchor.fraction,
				gesture.paper.getBoundingClientRect(),
				gesture.anchor.point,
			);
			viewport.scrollLeft += delta.x;
			viewport.scrollTop += delta.y;
			return;
		}
		// Each immediate step kept the scroll offsets, clamped to that step's range. The browser
		// clamps to the end zoom's range; only a gesture that went below both the start and the
		// end zoom was clamped further.
		const below = gesture.lowest < Math.min(gesture.base, gesture.target);
		const kept = keptScroll(gesture.scroll, gesture.layout, below ? gesture.lowest : undefined);
		viewport.scrollLeft = kept.x;
		viewport.scrollTop = kept.y;
	}
	/** Any other zoom (a command, the host) replaces the gesture. */
	wire(): () => void {
		const unsubscribe = this.controller.subscribe((state) => {
			if (this.#gesture && state.zoom !== this.#gesture.base) this.#end();
		});
		const Abort = this.viewport.ownerDocument.defaultView?.AbortController ?? AbortController;
		const events = new Abort();
		// Pointer input works on the rendered page, so it settles the zoom first.
		this.viewport.addEventListener('pointerdown', (event) => this.#press(event), {
			signal: events.signal,
			capture: true,
		});
		return () => {
			events.abort();
			unsubscribe();
			this.#end();
		};
	}
	/**
	 * Settles a pending zoom before a press reaches the page. Rendering can replace the element
	 * pressed (a selection handle), so the press then goes to the element now under the pointer:
	 * the preview showed the page where it settles, so that is the same handle or shape.
	 */
	#press(event: PointerEvent): void {
		if (!this.#gesture) return;
		const pressed = event.composedPath()[0] as Node | undefined;
		this.settle();
		if (!pressed || pressed.isConnected) return;
		const root = this.viewport.getRootNode() as Document | ShadowRoot;
		const now = root.elementFromPoint?.(event.clientX, event.clientY);
		const View = this.viewport.ownerDocument.defaultView;
		if (!now || !this.viewport.contains(now) || !View?.PointerEvent) return;
		event.stopImmediatePropagation();
		if (!now.dispatchEvent(new View.PointerEvent(event.type, event))) event.preventDefault();
	}
	#start(): Gesture | undefined {
		const paper = this.viewport.querySelector<SVGSVGElement>('svg.paper');
		const base = this.controller.state.zoom;
		const view = this.viewport.ownerDocument.defaultView;
		const box = paper?.getBoundingClientRect();
		if (!paper || !box?.width || !box.height || !view) return undefined;
		const port = this.viewport.getBoundingClientRect();
		const style = view.getComputedStyle(this.viewport);
		const px = (value: string) => Number.parseFloat(value) || 0;
		this.#gesture = {
			paper,
			base,
			layout: {
				origin: { x: port.left + this.viewport.clientLeft, y: port.top + this.viewport.clientTop },
				client: { x: this.viewport.clientWidth, y: this.viewport.clientHeight },
				padding: {
					left: px(style.paddingLeft),
					right: px(style.paddingRight),
					top: px(style.paddingTop),
					bottom: px(style.paddingBottom),
				},
				page: { x: box.width / base, y: box.height / base },
			},
			scroll: { x: this.viewport.scrollLeft, y: this.viewport.scrollTop },
			start: { x: box.left, y: box.top },
			target: base,
			lowest: base,
			anchor: SCROLL,
		};
		return this.#gesture;
	}
	#paint(): void {
		const gesture = this.#gesture;
		if (!gesture) return;
		const { layout, scroll, target, lowest, anchor, base } = gesture;
		const smallest = Math.min(lowest, target);
		const to = settledOrigin(
			layout,
			scroll,
			target,
			smallest < base ? smallest : undefined,
			anchor,
		);
		// The browser may clamp scrolling while the scaled paper is smaller than the laid-out one.
		const x = gesture.start.x + scroll.x - this.viewport.scrollLeft;
		const y = gesture.start.y + scroll.y - this.viewport.scrollTop;
		const style = gesture.paper.style;
		style.transformOrigin = '0 0';
		style.transform = `translate(${to.x - x}px, ${to.y - y}px) scale(${target / base})`;
		this.onPreview?.(target);
		// A slow frame must not let the timer run out between two inputs of the same gesture.
		this.#settleLater();
	}
	#end(): Gesture | undefined {
		const gesture = this.#gesture;
		this.#gesture = undefined;
		clearTimeout(this.#timer);
		this.#timer = undefined;
		if (this.#frame) this.viewport.ownerDocument.defaultView?.cancelAnimationFrame(this.#frame);
		this.#frame = 0;
		if (gesture) gesture.paper.style.transform = '';
		return gesture;
	}
}
