import { html } from 'lit';
import { OfficeElement, controlStyles } from '../base.js';
import { definer } from '../registry.js';
import { OFFICE_TOKENS } from '../tokens.js';
import css from './ruler.css?raw';

/** Tick subdivisions per unit for a given unit length in CSS pixels, finer as it grows. */
export function rulerDivisions(unitPixels: number): number {
	return unitPixels >= 192 ? 16 : unitPixels >= 96 ? 8 : unitPixels >= 48 ? 4 : 2;
}

/**
 * `<office-ui-ruler>`: a decorative Office ruler (inches by default). The product positions it
 * and says where zero is: `origin` is the CSS-pixel offset of zero from the ruler's start edge,
 * `scale` the CSS pixels per unit (96 at 100% for inches). `orientation="vertical"` draws a
 * side ruler; `direction="reverse"` makes values grow toward the start edge, as Visio's
 * vertical ruler grows upward from the page bottom. Hidden from assistive technology.
 */
export class OfficeUiRuler extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		orientation: { type: String, reflect: true },
		origin: { type: Number },
		scale: { type: Number },
		direction: { type: String },
	};
	declare orientation: string;
	declare origin: number;
	declare scale: number;
	declare direction: string;
	private frame = 0;
	private observer: ResizeObserver | undefined;

	constructor() {
		super();
		this.orientation = '';
		this.origin = 0;
		this.scale = 96;
		this.direction = '';
	}

	override connectedCallback(): void {
		super.connectedCallback();
		const Observer = this.ownerDocument.defaultView?.ResizeObserver;
		this.observer = Observer ? new Observer(() => this.schedule()) : undefined;
		this.observer?.observe(this);
		this.schedule();
	}

	override disconnectedCallback(): void {
		super.disconnectedCallback();
		this.observer?.disconnect();
		if (this.frame) this.ownerDocument.defaultView?.cancelAnimationFrame?.(this.frame);
		this.frame = 0;
	}

	private schedule(): void {
		if (this.frame || !this.isConnected) return;
		const view = this.ownerDocument.defaultView;
		this.frame = view?.requestAnimationFrame?.(() => this.draw()) ?? 0;
		if (!this.frame) this.draw();
	}

	protected override willUpdate(): void {
		this.setAttribute('aria-hidden', 'true');
	}

	protected override updated(): void {
		this.schedule();
	}

	private draw(): void {
		this.frame = 0;
		const canvas = this.renderRoot.querySelector('canvas');
		if (!canvas) return;
		const view = this.ownerDocument.defaultView;
		const vertical = this.orientation === 'vertical';
		const reverse = this.direction === 'reverse';
		const box = this.getBoundingClientRect();
		const length = vertical ? box.height : box.width;
		const thickness =
			(vertical ? box.width : box.height) || parseFloat(OFFICE_TOKENS['--office-ruler-size']);
		const ratio = view?.devicePixelRatio ?? 1;
		canvas.width = Math.max(1, Math.round((vertical ? thickness : length) * ratio));
		canvas.height = Math.max(1, Math.round((vertical ? length : thickness) * ratio));
		const context = canvas.getContext?.('2d');
		if (!context || !length) return;
		context.scale(ratio, ratio);
		// A canvas cannot read custom properties: draw with the host's computed, token-driven style.
		const style = view?.getComputedStyle(this);
		const ink = style?.color || OFFICE_TOKENS['--office-muted-foreground'];
		context.strokeStyle = ink;
		context.fillStyle = ink;
		context.lineWidth = 1;
		context.font = `${style?.fontSize || OFFICE_TOKENS['--office-font-size-3xs']} ${style?.fontFamily || OFFICE_TOKENS['--office-font']}`;
		const origin = Number.isFinite(this.origin) ? this.origin : 0;
		const unit = Math.max(1, Number.isFinite(this.scale) ? this.scale : 96);
		const divisions = rulerDivisions(unit);
		const step = unit / divisions;
		const sign = reverse ? -1 : 1;
		// Every tick index whose position falls inside the ruler, with one spare at each end.
		const a = (0 - origin) / (sign * step);
		const b = (length - origin) / (sign * step);
		const first = Math.floor(Math.min(a, b)) - 1;
		const last = Math.ceil(Math.max(a, b)) + 1;
		context.beginPath();
		for (let index = first; index <= last; index++) {
			const at = Math.round(origin + sign * index * step) + 0.5;
			const tick =
				index % divisions === 0
					? thickness
					: index % (divisions / 2) === 0
						? thickness * 0.5
						: thickness * 0.3;
			if (vertical) {
				context.moveTo(thickness, at);
				context.lineTo(thickness - tick, at);
			} else {
				context.moveTo(at, thickness);
				context.lineTo(at, thickness - tick);
			}
			if (index % divisions !== 0) continue;
			const label = String(index / divisions);
			if (!vertical) context.fillText(label, at + 2, 9);
			else {
				context.save();
				context.translate(9, at - 2);
				context.rotate(-Math.PI / 2);
				context.fillText(label, 0, 0);
				context.restore();
			}
		}
		context.stroke();
	}

	protected override render() {
		return html`<canvas></canvas>`;
	}
}

export const defineRuler = definer('office-ui-ruler', () => OfficeUiRuler);
