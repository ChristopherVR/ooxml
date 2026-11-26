import { html, nothing, svg } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { definer, emit } from '../registry';
import css from './ink-canvas.css?raw';

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
	};
	declare tool: InkCanvasTool;
	declare color: string;
	declare width: number;
	declare zoom: number;
	declare touch: boolean;
	declare label: string;
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
		if (!this.tool || this.#pointer !== undefined || event.button !== 0) return;
		if (event.pointerType === 'touch' && !this.touch) return;
		event.preventDefault();
		event.stopPropagation();
		this.#pointer = event.pointerId;
		this.#start = event.timeStamp;
		this.setPointerCapture?.(event.pointerId);
		this.#samples = [];
		this.#add(event);
	}

	#drag(event: PointerEvent): void {
		if (event.pointerId !== this.#pointer) return this.#track(event);
		event.preventDefault();
		const coalesced = event.getCoalescedEvents?.() ?? [];
		for (const sample of coalesced.length ? coalesced : [event]) this.#add(sample);
	}

	#add(event: PointerEvent): void {
		if (this.#samples.length >= MAX_SAMPLES) return;
		const point = this.#point(event);
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
			}
		</svg>`;
	}
}

export const defineInkCanvas = definer('office-ui-ink-canvas', () => OfficeUiInkCanvas);
