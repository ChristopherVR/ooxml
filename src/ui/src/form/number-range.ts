import { render } from 'lit';
import { rangeControl } from './range-control';

interface NumberRangeOptions {
	label(): string;
	enabled?(): boolean;
	onPreview?(value: number | undefined): void;
}

/** Pair a native range with a number field, committing once through the field's change handler. */
export function createNumberRange(input: HTMLInputElement, options: NumberRangeOptions) {
	const element = input.ownerDocument.createElement('div');
	let original: string | undefined;
	const enabled = () => !input.disabled && (options.enabled?.() ?? true);
	const cancel = () => {
		if (original === undefined) return;
		input.value = original;
		original = undefined;
		options.onPreview?.(undefined);
		paint();
	};
	const paint = () =>
		render(
			rangeControl({
				label: options.label(),
				value: input.valueAsNumber,
				valueText: input.value ? `${input.value}%` : '',
				min: Number(input.min),
				max: Number(input.max),
				step: Number(input.step) || 1,
				disabled: !enabled(),
				onInput: (value) => {
					if (!enabled() || !Number.isFinite(value)) return;
					original ??= input.value;
					input.value = String(value);
					options.onPreview?.(value);
					paint();
				},
				onChange: () => {
					if (original === undefined) return;
					if (!enabled()) return cancel();
					const value = input.value;
					cancel();
					input.value = value;
					input.dispatchEvent(new Event('change'));
					paint();
				},
			}),
			element,
		);
	element.addEventListener('pointercancel', cancel);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && original !== undefined) {
			event.preventDefault();
			event.stopPropagation();
			cancel();
		}
	});
	return {
		element,
		cancel,
		refresh: () => {
			cancel();
			paint();
		},
	};
}
