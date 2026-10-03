import { definer } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CSS = `
:host { display: block; box-sizing: border-box; background: var(--office-background, #fff);
	color: var(--office-muted-foreground, #605e5c); overflow: hidden; }
:host(:not([orientation="vertical"])) { height: 18px; border-bottom: 1px solid var(--office-border, #d1d5db); }
:host([orientation="vertical"]) { width: 18px; border-inline-end: 1px solid var(--office-border, #d1d5db); }
canvas { display: block; width: 100%; height: 100%; }
`;

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
export const defineRuler = definer('office-ui-ruler', () => {
	class OfficeUiRuler extends HTMLElement {
		static observedAttributes = ['orientation', 'origin', 'scale', 'direction'];
		readonly #canvas: HTMLCanvasElement;
		#frame = 0;
		#observer: ResizeObserver | undefined;
		constructor() {
			super();
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			this.#canvas = this.ownerDocument.createElement('canvas');
			root.append(this.#canvas);
		}
		connectedCallback(): void {
			this.setAttribute('aria-hidden', 'true');
			const Observer = this.ownerDocument.defaultView?.ResizeObserver;
			this.#observer = Observer ? new Observer(() => this.#schedule()) : undefined;
			this.#observer?.observe(this);
			this.#schedule();
		}
		disconnectedCallback(): void {
			this.#observer?.disconnect();
			if (this.#frame) this.ownerDocument.defaultView?.cancelAnimationFrame?.(this.#frame);
			this.#frame = 0;
		}
		attributeChangedCallback(): void {
			this.#schedule();
		}
		#number(name: string, fallback: number): number {
			const value = Number(this.getAttribute(name));
			return this.hasAttribute(name) && Number.isFinite(value) ? value : fallback;
		}
		#schedule(): void {
			if (this.#frame || !this.isConnected) return;
			const view = this.ownerDocument.defaultView;
			this.#frame = view?.requestAnimationFrame?.(() => this.#draw()) ?? 0;
			if (!this.#frame) this.#draw();
		}
		#draw(): void {
			this.#frame = 0;
			const view = this.ownerDocument.defaultView;
			const vertical = this.getAttribute('orientation') === 'vertical';
			const reverse = this.getAttribute('direction') === 'reverse';
			const box = this.getBoundingClientRect();
			const length = vertical ? box.height : box.width;
			const thickness = (vertical ? box.width : box.height) || 18;
			const ratio = view?.devicePixelRatio ?? 1;
			const canvas = this.#canvas;
			canvas.width = Math.max(1, Math.round((vertical ? thickness : length) * ratio));
			canvas.height = Math.max(1, Math.round((vertical ? length : thickness) * ratio));
			const context = canvas.getContext?.('2d');
			if (!context || !length) return;
			context.scale(ratio, ratio);
			const ink = view?.getComputedStyle(this).color || '#605e5c';
			context.strokeStyle = ink;
			context.fillStyle = ink;
			context.lineWidth = 1;
			context.font = `9px ${view?.getComputedStyle(this).fontFamily || 'system-ui, sans-serif'}`;
			const origin = this.#number('origin', 0);
			const unit = Math.max(1, this.#number('scale', 96));
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
	}
	return OfficeUiRuler;
});
