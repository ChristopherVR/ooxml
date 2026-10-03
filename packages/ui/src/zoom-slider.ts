import { tok } from './tokens.js';
import { createIconSvg, paintIcon } from './icons.js';
import { definer, emit } from './registry.js';
import { attachStyles, controlCss } from './styles.js';

const CSS = `
:host { display: inline-flex; align-items: center; gap: ${tok('--office-space-0')}; font-size: ${tok('--office-font-size-xs')};
	color: ${tok('--office-foreground')}; font-family: ${tok('--office-font')}; }
button { box-sizing: border-box; display: inline-grid; place-items: center; min-width: ${tok('--office-target-size')};
	min-height: ${tok('--office-target-size')}; padding: 0; border: ${tok('--office-border-width')} solid transparent;
	border-radius: ${tok('--office-radius')}; background: transparent; color: inherit; font: inherit;
	font-size: ${tok('--office-font-size-lg')}; line-height: 1; cursor: pointer; }
button:hover:not(:disabled) { background: ${tok('--office-background')}; }
button:focus-visible, input:focus-visible { outline: ${tok('--office-focus-width')} solid ${tok('--office-ring')}; outline-offset: calc(${tok('--office-focus-offset')} / 2); }
button:disabled, input:disabled { opacity: .5; cursor: not-allowed; }
input { width: ${tok('--office-zoom-slider-width')}; min-width: ${tok('--office-zoom-slider-min-width')}; margin: 0 ${tok('--office-space-0')};
	accent-color: ${tok('--office-accent')}; }
output { min-width: ${tok('--office-zoom-value-width')}; text-align: center; font-variant-numeric: tabular-nums; }
svg { width: ${tok('--office-icon-size-md')}; height: ${tok('--office-icon-size-md')}; fill: none; stroke: currentColor; stroke-width: ${tok('--office-icon-stroke')};
	stroke-linecap: round; stroke-linejoin: round; }
:host(:not([fit])) .fit { display: none; }
@media (pointer: coarse), (max-width: 767px) { input { flex: 1; } }
@media (forced-colors: active) { button { color: ButtonText; } button:disabled { color: GrayText; } }
`;

const clampTo = (value: number, min: number, max: number) =>
	Math.min(max, Math.max(min, Math.round(value)));
const number = (value: string | null, fallback: number) => {
	const parsed = Number(value);
	return value !== null && Number.isFinite(parsed) ? parsed : fallback;
};

/**
 * Office status-bar zoom: zoom out, slider, zoom in, percentage and an optional fit button.
 * Attributes: `value` (percent, default 100), `min` (10), `max` (400), `step` (10, for the
 * buttons), `disabled`, `fit` (shows the fit button; its value is the button's name, e.g.
 * "Fit page to current window"), `label` (group name, default "Zoom").
 * User changes emit native `input` (slider drag) and `change` events with `value` updated;
 * the fit button emits `office-command` `{ command: 'zoom-fit' }`. Setting `value` never emits.
 */
export const defineZoomSlider = definer('office-ui-zoom-slider', () => {
	class OfficeUiZoomSlider extends HTMLElement {
		static observedAttributes = ['value', 'min', 'max', 'disabled', 'fit', 'label'];
		private readonly out: HTMLButtonElement;
		private readonly in: HTMLButtonElement;
		private readonly range: HTMLInputElement;
		private readonly output: HTMLOutputElement;
		private readonly fit: HTMLButtonElement;
		constructor() {
			super();
			const doc = this.ownerDocument;
			const root = this.attachShadow({ mode: 'open' });
			attachStyles(root, controlCss(CSS));
			const button = (name: string, text: string) => {
				const element = doc.createElement('button');
				element.type = 'button';
				element.setAttribute('aria-label', name);
				element.title = name;
				element.textContent = text;
				return element;
			};
			this.out = button('Zoom out', '−');
			this.in = button('Zoom in', '+');
			this.range = doc.createElement('input');
			this.range.type = 'range';
			this.range.setAttribute('aria-label', 'Zoom');
			this.output = doc.createElement('output');
			this.fit = button('Fit', '');
			this.fit.className = 'fit';
			const svg = createIconSvg(doc);
			paintIcon(svg, 'fitPage');
			this.fit.append(svg);
			root.append(this.out, this.range, this.in, this.output, this.fit);
			this.out.addEventListener('click', () => this.stepBy(-1));
			this.in.addEventListener('click', () => this.stepBy(1));
			this.range.addEventListener('input', () => this.commit(this.range.valueAsNumber, 'input'));
			this.range.addEventListener('change', () =>
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true })),
			);
			this.fit.addEventListener('click', () =>
				emit(this, 'office-command', { command: 'zoom-fit' }),
			);
		}
		connectedCallback(): void {
			this.setAttribute('role', 'group');
			this.sync();
		}
		attributeChangedCallback(): void {
			this.sync();
		}
		get min(): number {
			return number(this.getAttribute('min'), 10);
		}
		get max(): number {
			return Math.max(this.min, number(this.getAttribute('max'), 400));
		}
		get value(): number {
			return clampTo(number(this.getAttribute('value'), 100), this.min, this.max);
		}
		set value(next: number) {
			this.setAttribute('value', String(clampTo(Number(next) || 0, this.min, this.max)));
		}
		get disabled(): boolean {
			return this.hasAttribute('disabled');
		}
		set disabled(value: boolean) {
			this.toggleAttribute('disabled', Boolean(value));
		}
		/** Office buttons snap to the next multiple of `step`, so 67% steps to 70%, not 77%. */
		private stepBy(direction: 1 | -1): void {
			const step = Math.max(1, number(this.getAttribute('step'), 10));
			const current = this.value;
			const next =
				direction > 0
					? Math.floor(current / step) * step + step
					: Math.ceil(current / step) * step - step;
			this.commit(next, 'change');
		}
		private commit(next: number, type: 'input' | 'change'): void {
			if (this.disabled) return;
			const value = clampTo(next, this.min, this.max);
			if (value === this.value && type === 'change') return;
			this.value = value;
			this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
			if (type === 'change')
				this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
		}
		private sync(): void {
			const { min, max, value, disabled } = this;
			this.setAttribute('aria-label', this.getAttribute('label') ?? 'Zoom');
			this.range.min = String(min);
			this.range.max = String(max);
			this.range.value = String(value);
			this.range.setAttribute('aria-valuetext', `${value}%`);
			this.output.value = `${value}%`;
			this.out.disabled = disabled || value <= min;
			this.in.disabled = disabled || value >= max;
			this.range.disabled = disabled;
			const fit = this.getAttribute('fit');
			this.fit.disabled = disabled;
			this.fit.setAttribute('aria-label', fit || 'Fit');
			this.fit.title = fit || 'Fit';
		}
	}
	return OfficeUiZoomSlider;
});
