import { html } from 'lit';
import { OfficeElement, controlStyles, flag } from '../base';
import { glyph } from '../glyph';
import { definer, present } from '../registry';
import css from './zoom-slider.css?raw';

const clampTo = (value: number, min: number, max: number) =>
	Math.min(max, Math.max(min, Math.round(value)));
const finite = (value: number, fallback: number) => (Number.isFinite(value) ? value : fallback);

/**
 * Office status-bar zoom: zoom out, slider, zoom in, percentage and an optional fit button.
 * Attributes: `value` (percent, default 100), `min` (10), `max` (400), `step` (10, for the
 * buttons), `disabled`, `fit` (shows the fit button; its value is the button's name, e.g.
 * "Fit page to current window"), `label` (group name, default "Zoom"). `out-label`, `in-label` and
 * `slider-label` name the zoom-out button, zoom-in button and slider (English by default), so a
 * product can translate them.
 * User changes emit native `input` (slider drag) and `change` events with `value` updated;
 * the fit button emits `office-command` `{ command: 'zoom-fit' }`. Setting `value` never emits.
 */
export class OfficeUiZoomSlider extends OfficeElement {
	static override styles = controlStyles(css);
	static override properties = {
		// `value` clamps to the range, so it has a hand-written accessor below.
		value: { type: Number, reflect: true, noAccessor: true },
		min: { type: Number },
		max: { type: Number },
		step: { type: Number },
		disabled: flag,
		fit: { type: String, reflect: true },
		label: { type: String },
		outLabel: { type: String, attribute: 'out-label' },
		inLabel: { type: String, attribute: 'in-label' },
		sliderLabel: { type: String, attribute: 'slider-label' },
	};
	declare min: number;
	declare max: number;
	declare step: number;
	declare disabled: boolean;
	declare fit: string | null;
	declare label: string | null;
	declare outLabel: string | null;
	declare inLabel: string | null;
	declare sliderLabel: string | null;
	private requested = 100;

	constructor() {
		super();
		this.min = 10;
		this.max = 400;
		this.step = 10;
		this.disabled = false;
		this.fit = null;
		this.label = null;
		this.outLabel = null;
		this.inLabel = null;
		this.sliderLabel = null;
	}

	get value(): number {
		return clampTo(this.requested, this.low, this.high);
	}
	set value(next: number) {
		const old = this.value;
		this.requested = finite(Number(next), 100);
		this.requestUpdate('value', old);
	}

	private get low(): number {
		return finite(this.min, 10);
	}
	private get high(): number {
		return Math.max(this.low, finite(this.max, 400));
	}

	/** Office buttons snap to the next multiple of `step`, so 67% steps to 70%, not 77%. */
	private stepBy(direction: 1 | -1): void {
		const step = Math.max(1, finite(this.step, 10));
		const current = this.value;
		const next =
			direction > 0
				? Math.floor(current / step) * step + step
				: Math.ceil(current / step) * step - step;
		this.commit(next, 'change');
	}

	private commit(next: number, type: 'input' | 'change'): void {
		if (present(this.disabled)) return;
		const value = clampTo(next, this.low, this.high);
		if (value === this.value && type === 'change') return;
		this.value = value;
		this.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
		if (type === 'change')
			this.dispatchEvent(new Event('change', { bubbles: true, composed: true }));
	}

	protected override willUpdate(): void {
		this.setAttribute('role', 'group');
		this.setAttribute('aria-label', this.label ?? 'Zoom');
	}

	protected override render() {
		const { value } = this;
		const disabled = present(this.disabled);
		const fit = this.fit || 'Fit';
		const out = this.outLabel || 'Zoom out';
		const inn = this.inLabel || 'Zoom in';
		return html`
			<button
				type="button"
				aria-label=${out}
				title=${out}
				?disabled=${disabled || value <= this.low}
				@click=${() => this.stepBy(-1)}
				>−</button
			>
			<input
				type="range"
				aria-label=${this.sliderLabel || 'Zoom'}
				aria-valuetext="${value}%"
				min=${this.low}
				max=${this.high}
				.value=${String(value)}
				?disabled=${disabled}
				@input=${(event: Event) =>
					this.commit((event.target as HTMLInputElement).valueAsNumber, 'input')}
				@change=${() => this.dispatchEvent(new Event('change', { bubbles: true, composed: true }))}
			/>
			<button
				type="button"
				aria-label=${inn}
				title=${inn}
				?disabled=${disabled || value >= this.high}
				@click=${() => this.stepBy(1)}
				>+</button
			>
			<output>${value}%</output>
			<button
				class="fit"
				type="button"
				aria-label=${fit}
				title=${fit}
				?disabled=${disabled}
				@click=${() => this.fire('office-command', { command: 'zoom-fit' })}
				>${glyph('fitPage', 'fit-icon')}</button
			>
		`;
	}
}

export const defineZoomSlider = definer('office-ui-zoom-slider', () => OfficeUiZoomSlider);
