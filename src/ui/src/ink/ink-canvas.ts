import { html, nothing, svg } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, emit } from '../registry';
import css from './ink-canvas.css?raw';
import { InkRuler } from './ink-ruler';

/** The Draw tool the canvas captures for; anything else lets the pointer through. */
export type InkCanvasTool = '' | 'pen' | 'pencil' | 'highlighter' | 'eraser' | 'lasso';

/** One pointer sample, in CSS pixels of the canvas at 100% zoom. */
export interface InkCanvasSample {
	x: number;
	y: number;
	/** 0..1, as the pointer reports it (a mouse reports a constant). */
	pressure: number;
	/** Milliseconds since the stroke began. */
	time: number;
}

/** `office-ink-stroke`: a pen, pencil or highlighter stroke the user finished. */
export interface InkStrokeDetail {
	tool: 'pen' | 'pencil' | 'highlighter';
	samples: InkCanvasSample[];
}

/** `office-ink-erase`: the eraser touched this point (fired on press and every move). */
export interface InkEraseDetail {
	x: number;
	y: number;
	radius: number;
	/** True on the last event of one erasing gesture. */
	done: boolean;
}

/** `office-ink-lasso`: the closed outline the lasso drew. */
export interface InkLassoDetail {
	points: { x: number; y: number }[];
}

const COLOR = /^#[\da-f]{6}$/iu;
/** The eraser's reach, in CSS pixels at 100% zoom. */
export const INK_ERASER_RADIUS = 6;
/** Samples beyond this are dropped: a stroke this long is a runaway gesture. */
const MAX_SAMPLES = 4096;

/**
 * `<office-ui-ink-canvas>`: the Draw tab's pointer surface. The product lays it over the surface it
 * draws on (a page, a sheet) and sets `tool`, the pen's `color` and `width`, and the surface's
 * `zoom`. While a tool is on it captures the pointer and draws the stroke being made, then fires
 * `office-ink-stroke`, `office-ink-erase` or `office-ink-lasso` with coordinates at 100% zoom; the
 * product turns them into document ink. With no tool (the default) it is hidden and lets the
 * pointer through. Touch scrolls unless `touch` (Draw with Touch) is set; a pen or mouse always
 * draws. Escape abandons the gesture in progress (`office-ink-cancel`).
 *
 * `ruler` lays the ruler stencil over the surface, with or without a tool: it is dragged, rotated
 * with the mouse wheel or Shift and the arrow keys (arrows alone move it), and a pen stroke begun
 * near its edge follows the edge in a straight line. `locale` names it in the product's language.
 */
export class OfficeUiInkCanvas extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		tool: { type: String, reflect: true },
		color: { type: String },
		width: { type: Number },
		zoom: { type: Number },
		touch: { ...flag, reflect: true },
		label: { type: String },
		ruler: { ...flag, reflect: true },
		locale: { type: String },
	};
	declare tool: InkCanvasTool;
	declare color: string;
	declare width: number;
	declare zoom: number;
	declare touch: boolean;
	declare label: string;
	declare ruler: boolean;
	declare locale: string;
	readonly #ruler = new InkRuler();
	/** The ruler being dragged: its pointer and where it was grabbed from its centre. */
	#rulerDrag: { id: number; dx: number; dy: number } | undefined;
	/** The stroke in progress follows the ruler's edge. */
	#snap = false;
	#samples: InkCanvasSample[] = [];
	#pointer: number | undefined;
	#start = 0;
	#cursor: { x: number; y: number } | undefined;
	readonly #down = (event: PointerEvent) => this.#press(event);
	readonly #move = (event: PointerEvent) => this.#drag(event);
	readonly #up = (event: PointerEvent) => this.#release(event, false);
	readonly #lost = (event: PointerEvent) => this.#release(event, true);
	readonly #hover = (event: PointerEvent) => this.#track(event);
	readonly #leave = () => this.#untrack();
	readonly #key = (event: KeyboardEvent) => {
		if (event.key === 'Escape' && this.#pointer !== undefined) this.cancel();
	};

	constructor() {
		super();
		this.tool = '';
		this.color = '#000000';
		this.width = 2;
		this.zoom = 1;
		this.touch = false;
		this.label = '';
		this.ruler = false;
		this.locale = '';
	}

	/** The ruler's angle in degrees (-180..180), as the user turned it. */
	get rulerAngle(): number {
		return this.#ruler.angle;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		this.addEventListener('pointerdown', this.#down);
		this.addEventListener('pointermove', this.#move);
		this.addEventListener('pointerup', this.#up);
		this.addEventListener('pointercancel', this.#lost);
		this.addEventListener('lostpointercapture', this.#lost);
		this.addEventListener('pointerover', this.#hover);
		this.addEventListener('pointerleave', this.#leave);
		this.ownerDocument.addEventListener('keydown', this.#key, true);
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.removeEventListener('pointerdown', this.#down);
		this.removeEventListener('pointermove', this.#move);
		this.removeEventListener('pointerup', this.#up);
		this.removeEventListener('pointercancel', this.#lost);
		this.removeEventListener('lostpointercapture', this.#lost);
		this.removeEventListener('pointerover', this.#hover);
		this.removeEventListener('pointerleave', this.#leave);
		this.ownerDocument.removeEventListener('keydown', this.#key, true);
		this.#reset();
	}

	protected override updated(): void {
		if (this.ruler && !this.#ruler.center) this.#placeRuler();
	}

	/** The ruler goes in the middle of what is visible of the surface, the first time it shows. */
	#placeRuler(): void {
		const rect = this.getBoundingClientRect();
		const clip = this.parentElement?.getBoundingClientRect();
		const view = this.ownerDocument.defaultView;
		const left = Math.max(rect.left, clip?.left ?? -Infinity, 0);
		const top = Math.max(rect.top, clip?.top ?? -Infinity, 0);
		const right = Math.min(rect.right, clip?.right ?? Infinity, view?.innerWidth ?? Infinity);
		const bottom = Math.min(rect.bottom, clip?.bottom ?? Infinity, view?.innerHeight ?? Infinity);
		const scale = this.#scale();
		const visible =
			right > left && bottom > top
				? {
						x: (left - rect.left) / scale,
						y: (top - rect.top) / scale,
						width: (right - left) / scale,
						height: (bottom - top) / scale,
					}
				: { x: 0, y: 0, width: 600, height: 300 };
		this.#ruler.place(visible);
		this.requestUpdate();
	}

	#onRuler(event: Event): boolean {
		return event
			.composedPath()
			.some((node) => node instanceof Element && node.classList.contains('ruler'));
	}

	#grabRuler(event: PointerEvent): void {
		event.preventDefault();
		event.stopPropagation();
		const point = this.#point(event);
		const center = this.#ruler.center ?? point;
		this.#rulerDrag = { id: event.pointerId, dx: center.x - point.x, dy: center.y - point.y };
		const ruler = this.renderRoot.querySelector<SVGGElement>('.ruler');
		ruler?.focus({ preventScroll: true });
		ruler?.setPointerCapture?.(event.pointerId);
	}

	readonly #rulerKey = (event: KeyboardEvent): void => {
		if (!this.#ruler.key(event)) return;
		event.preventDefault();
		event.stopPropagation();
		this.requestUpdate();
	};

	readonly #rulerWheel = (event: WheelEvent): void => {
		event.preventDefault();
		event.stopPropagation();
		this.#ruler.wheel(event);
		this.requestUpdate();
	};

	protected override willUpdate(): void {
		this.hostWrite(() => {
			if (this.label) this.setAttribute('aria-label', this.label);
			else this.removeAttribute('aria-label');
			if (this.tool) this.setAttribute('role', 'application');
			else this.removeAttribute('role');
		});
	}

	/** Abandons the gesture in progress. */
	cancel(): void {
		if (this.#pointer === undefined) return;
		this.#reset();
		emit(this, 'office-ink-cancel', {});
	}

	#scale(): number {
		return Number.isFinite(this.zoom) && this.zoom > 0 ? this.zoom : 1;
	}

	#point(event: PointerEvent): { x: number; y: number } {
		const rect = this.getBoundingClientRect();
		const scale = this.#scale();
		return { x: (event.clientX - rect.left) / scale, y: (event.clientY - rect.top) / scale };
	}

	#press(event: PointerEvent): void {
		if (this.#pointer !== undefined || this.#rulerDrag || event.button !== 0) return;
		// The ruler takes a finger too, as Office's does.
		if (this.ruler && this.#onRuler(event)) return this.#grabRuler(event);
		if (!this.tool) return;
		if (event.pointerType === 'touch' && !this.touch) return;
		event.preventDefault();
		event.stopPropagation();
		this.#pointer = event.pointerId;
		this.#start = event.timeStamp;
		this.setPointerCapture?.(event.pointerId);
		this.#samples = [];
		const tool = this.tool;
		this.#snap =
			this.ruler &&
			(tool === 'pen' || tool === 'pencil' || tool === 'highlighter') &&
			this.#ruler.near(this.#point(event));
		this.#add(event);
	}

	#drag(event: PointerEvent): void {
		const grab = this.#rulerDrag;
		if (grab && event.pointerId === grab.id) {
			event.preventDefault();
			const point = this.#point(event);
			this.#ruler.moveTo({ x: point.x + grab.dx, y: point.y + grab.dy });
			this.requestUpdate();
			return;
		}
		if (event.pointerId !== this.#pointer) return this.#track(event);
		event.preventDefault();
		const coalesced = event.getCoalescedEvents?.() ?? [];
		for (const sample of coalesced.length ? coalesced : [event]) this.#add(sample);
	}

	#add(event: PointerEvent): void {
		if (this.#samples.length >= MAX_SAMPLES) return;
		const point = this.#snap ? this.#ruler.snap(this.#point(event)) : this.#point(event);
		const pressure = event.pointerType === 'mouse' ? 0.5 : event.pressure || 0.5;
		this.#samples.push({ ...point, pressure, time: Math.max(0, event.timeStamp - this.#start) });
		if (this.tool === 'eraser') {
			this.#cursor = point;
			emit<InkEraseDetail>(this, 'office-ink-erase', {
				...point,
				radius: INK_ERASER_RADIUS,
				done: false,
			});
		}
		this.requestUpdate();
	}

	#release(event: PointerEvent, lost: boolean): void {
		if (this.#rulerDrag?.id === event.pointerId) {
			this.#rulerDrag = undefined;
			return;
		}
		if (event.pointerId !== this.#pointer) return;
		const samples = this.#samples;
		const tool = this.tool;
		this.#pointer = undefined;
		if (this.hasPointerCapture?.(event.pointerId)) this.releasePointerCapture(event.pointerId);
		this.#samples = [];
		this.requestUpdate();
		if (lost && event.type === 'pointercancel') {
			emit(this, 'office-ink-cancel', {});
			return;
		}
		if (tool === 'pen' || tool === 'pencil' || tool === 'highlighter')
			emit<InkStrokeDetail>(this, 'office-ink-stroke', { tool, samples });
		else if (tool === 'lasso' && samples.length > 2)
			emit<InkLassoDetail>(this, 'office-ink-lasso', {
				points: samples.map(({ x, y }) => ({ x, y })),
			});
		else if (tool === 'eraser') {
			const last = samples.at(-1);
			if (last)
				emit<InkEraseDetail>(this, 'office-ink-erase', {
					x: last.x,
					y: last.y,
					radius: INK_ERASER_RADIUS,
					done: true,
				});
		}
	}

	#track(event: PointerEvent): void {
		if (this.tool !== 'eraser' || this.#pointer !== undefined) return;
		this.#cursor = this.#point(event);
		this.requestUpdate();
	}

	#untrack(): void {
		if (!this.#cursor || this.#pointer !== undefined) return;
		this.#cursor = undefined;
		this.requestUpdate();
	}

	#reset(): void {
		this.#pointer = undefined;
		this.#rulerDrag = undefined;
		this.#snap = false;
		this.#samples = [];
		this.#cursor = undefined;
		this.requestUpdate();
	}

	protected override render() {
		const scale = this.#scale();
		const points = this.#samples.map(({ x, y }) => `${x * scale},${y * scale}`).join(' ');
		const tool = this.tool;
		const ink = tool === 'pen' || tool === 'pencil' || tool === 'highlighter';
		const color = COLOR.test(this.color) ? this.color : '#000000';
		const width = Math.max(0.5, Number(this.width) || 2) * scale;
		const opacity = tool === 'highlighter' ? 0.4 : tool === 'pencil' ? 0.72 : 1;
		const cursor = tool === 'eraser' ? this.#cursor : undefined;
		return html`<svg aria-hidden="true">
				${
					ink && points
						? svg`<polyline class="live" points=${points} stroke=${color} stroke-width=${width} stroke-opacity=${opacity}></polyline>`
						: nothing
				}${
					tool === 'lasso' && points
						? svg`<polyline class="lasso" points=${points}></polyline>`
						: nothing
				}${
					cursor
						? svg`<circle class="eraser" cx=${cursor.x * scale} cy=${cursor.y * scale} r=${INK_ERASER_RADIUS * scale}></circle>`
						: nothing
				}</svg
			>${
				// Its own layer: the stroke being drawn is hidden from assistive technology, the ruler is not.
				this.ruler && this.#ruler.center
					? html`<svg class="ruler-layer">
							${this.#ruler.render(scale, this.locale || undefined, {
								keydown: this.#rulerKey,
								wheel: this.#rulerWheel,
							})}
						</svg>`
					: nothing
			}`;
	}
}

export const defineInkCanvas = definer('office-ui-ink-canvas', () => OfficeUiInkCanvas);
