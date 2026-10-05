// Zoom: preset magnifications, Fit selection or a custom percentage (10 to 400).
import { currentZoom, setZoom, zoomToFit } from '../commands/view.js';
import type { EditorContext } from '../context.js';
import { field, numberInput, radios } from './fields.js';
import { showDialog } from './frame.js';

const CHOICES: ReadonlyArray<readonly [string, string]> = [
	['200', '200%'],
	['100', '100%'],
	['75', '75%'],
	['50', '50%'],
	['25', '25%'],
	['fit', 'Fit selection'],
	['custom', 'Custom:'],
];

export function openZoom(ctx: EditorContext): Promise<number | undefined> {
	const zoom = currentZoom(ctx);
	const preset = CHOICES.find(([v]) => v === String(zoom))?.[0] ?? 'custom';
	const choice = radios(ctx, 'Magnification', CHOICES, preset);
	const custom = numberInput(ctx, zoom, 10, 400);
	custom.addEventListener('input', () => choice.set('custom'));
	return showDialog<number>(ctx, {
		name: 'zoom',
		heading: 'Zoom',
		body: [choice.element, field(ctx, 'Percent', custom)],
		opened: () => (choice.inputs.find((i) => i.checked) ?? custom).focus(),
		submit: () => {
			const value = choice.get();
			let percent: number | undefined;
			if (value === 'fit') percent = zoomToFit(ctx);
			else if (value === 'custom') {
				const n = Number(custom.value);
				if (!custom.value.trim() || !Number.isFinite(n) || n < 10 || n > 400) {
					ctx.toast(
						ctx.t('Enter a number between {min} and {max}.', { min: 10, max: 400 }),
						'warning',
					);
					custom.focus();
					return undefined;
				}
				percent = Math.round(n);
			} else percent = Number(value);
			if (percent === undefined) return undefined;
			setZoom(ctx, percent);
			return percent;
		},
	});
}
