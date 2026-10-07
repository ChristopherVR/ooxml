import { html } from 'lit';

export interface RangeControlOptions {
	label: string;
	value: number;
	valueText: string;
	min: number;
	max: number;
	step?: number;
	disabled: boolean;
	onInput(value: number): void;
	onChange(): void;
}

/** Native keyboard, pointer and accessibility behavior shared by Office range controls. */
export function rangeControl(options: RangeControlOptions) {
	return html`<input
		class="office-range-control"
		type="range"
		aria-label=${options.label}
		aria-valuetext=${options.valueText}
		min=${options.min}
		max=${options.max}
		step=${options.step ?? 1}
		.value=${String(options.value)}
		?disabled=${options.disabled}
		@input=${(event: Event) => {
			if (!options.disabled) options.onInput((event.target as HTMLInputElement).valueAsNumber);
		}}
		@change=${() => {
			if (!options.disabled) options.onChange();
		}}
	/>`;
}
