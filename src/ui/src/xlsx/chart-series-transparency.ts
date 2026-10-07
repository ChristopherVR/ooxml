import { render } from 'lit';
import {
	chartSeriesTransparency,
	chartSeriesTransparencyPatch,
	type ChartObject,
} from 'ooxml-core/xlsx';
import { activeChart, type EditorContext } from 'ooxml-core/xlsx/ui';
import { rangeControl } from '../form/range-control';
import { el, field, numberInput } from './dialogs/fields';

/** The shared native range input pairs with the existing validated percentage field. */
export function createSeriesTransparency(ctx: EditorContext, selected: () => number) {
	const element = el(ctx, 'div');
	const input = numberInput(ctx, 0, 0, 100);
	const row = field(ctx, 'Transparency', input);
	const unit = el(ctx, 'span');
	unit.textContent = '%';
	row.append(unit);
	const host = el(ctx, 'div', 'xve-chart-series-range');
	element.append(row, host);
	let current: ChartObject | undefined;
	const paint = () =>
		render(
			rangeControl({
				label: ctx.t('Transparency'),
				value: input.valueAsNumber,
				valueText: `${input.value}%`,
				min: 0,
				max: 100,
				disabled: input.disabled,
				onInput: (value) => {
					if (!ctx.commands.isEnabled('chart.format-series')) return;
					input.value = String(value);
					paint();
				},
				onChange: () => input.dispatchEvent(new Event('change')),
			}),
			host,
		);
	const refresh = (chart: ChartObject | undefined) => {
		current = chart;
		const value = chart ? chartSeriesTransparency(chart, selected()) : undefined;
		input.value = String(value ?? 0);
		input.disabled = value === undefined || !ctx.commands.isEnabled('chart.format-series');
		row.querySelector('span')!.textContent = ctx.t('Transparency');
		input.setAttribute('aria-label', ctx.t('Transparency'));
		paint();
	};
	input.addEventListener('change', () => {
		if (!current || input.disabled || !ctx.commands.isEnabled('chart.format-series'))
			return refresh(current);
		if (!input.checkValidity() || !Number.isInteger(input.valueAsNumber)) {
			ctx.toast(
				ctx.t('Enter a whole number from {min} to {max}.', { min: 0, max: 100 }),
				'warning',
			);
			return refresh(current);
		}
		const found = activeChart(ctx);
		if (!found || found.chart !== current) return refresh(current);
		const patch = chartSeriesTransparencyPatch(current, selected(), input.valueAsNumber);
		if (patch) ctx.session()?.updateChart(ctx.activeSheet(), found.index, patch);
		refresh(found.chart);
	});
	return { element, refresh };
}
