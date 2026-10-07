// Chart group spacing lives in core; this docked pane only binds the native controls to edits.
import { chartBarSpacing } from 'ooxml-core/xlsx';
import {
	activeChart,
	editing,
	editChart,
	icon,
	type Command,
	type EditorContext,
} from 'ooxml-core/xlsx/ui';
import { el, field, numberInput } from './dialogs/fields';
import { render } from 'lit';
import { rangeControl } from '../form/range-control';

export function chartSeriesCommand(): Command {
	return editing({
		id: 'chart.format-series',
		label: 'Format Data Series',
		icon: icon('chartColumn'),
		lock: 'objects',
		enabled: (ctx) => {
			const type = activeChart(ctx)?.chart.chartType;
			return type === 'column' || type === 'bar';
		},
		run: (ctx) => void ctx.dialogs.open('format-chart-series'),
	});
}

export function createChartSeriesPane(ctx: EditorContext) {
	const element = el(ctx, 'aside', 'xve-chart-series-pane');
	element.hidden = true;
	element.setAttribute('role', 'complementary');
	const header = el(ctx, 'div', 'xve-chart-series-header');
	const title = el(ctx, 'h2');
	const closeButton = el(ctx, 'button', 'xve-icon-button');
	closeButton.type = 'button';
	closeButton.textContent = '×';
	const close = () => {
		element.hidden = true;
		ctx.grid()?.focus();
	};
	closeButton.addEventListener('click', close);
	const body = el(ctx, 'div', 'xve-chart-series-body');
	const section = el(ctx, 'h3');
	const overlap = numberInput(ctx, 0, -100, 100);
	const gap = numberInput(ctx, 150, 0, 500);
	const overlapField = field(ctx, 'Series Overlap', overlap);
	const gapField = field(ctx, 'Gap Width', gap);
	const overlapRange = el(ctx, 'div', 'xve-chart-series-range');
	const gapRange = el(ctx, 'div', 'xve-chart-series-range');
	const ranges = [
		[overlapRange, overlap, 'Series Overlap'],
		[gapRange, gap, 'Gap Width'],
	] as const;
	const paintRange = (host: HTMLElement, input: HTMLInputElement, key: string) => {
		const value = input.valueAsNumber;
		render(
			rangeControl({
				label: ctx.t(key),
				value,
				valueText: `${value}%`,
				min: Number(input.min),
				max: Number(input.max),
				disabled: input.disabled,
				onInput: (next) => {
					if (!ctx.commands.isEnabled('chart.format-series')) return refresh();
					input.value = String(next);
					paintRange(host, input, key);
				},
				onChange: () => input.dispatchEvent(new Event('change')),
			}),
			host,
		);
	};
	for (const row of [overlapField, gapField]) {
		const unit = el(ctx, 'span');
		unit.textContent = '%';
		row.append(unit);
	}
	const empty = el(ctx, 'p', 'xve-note');
	header.append(title, closeButton);
	body.append(section, overlapField, overlapRange, gapField, gapRange);
	element.append(header, body, empty);
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Escape') return;
		event.preventDefault();
		event.stopPropagation();
		close();
	});
	const relocalize = () => {
		title.textContent = ctx.t('Format Data Series');
		element.setAttribute('aria-label', title.textContent);
		closeButton.setAttribute('aria-label', ctx.t('Close'));
		section.textContent = ctx.t('Series Options');
		empty.textContent = ctx.t('Select a bar or column chart to format its series.');
		for (const [row, input, key] of [
			[overlapField, overlap, 'Series Overlap'],
			[gapField, gap, 'Gap Width'],
		] as const) {
			row.querySelector('span')!.textContent = ctx.t(key);
			input.setAttribute('aria-label', ctx.t(key));
		}
		for (const [host, input, key] of ranges) paintRange(host, input, key);
	};
	const refresh = () => {
		if (element.hidden) return;
		const chart = activeChart(ctx)?.chart;
		const supported = chart?.chartType === 'bar' || chart?.chartType === 'column';
		body.hidden = !supported;
		empty.hidden = supported;
		const spacing = chartBarSpacing(chart ?? {});
		gap.value = String(spacing.barGapWidth);
		overlap.value = String(spacing.barOverlap);
		gap.disabled = overlap.disabled = !ctx.commands.isEnabled('chart.format-series');
		for (const [host, input, key] of ranges) paintRange(host, input, key);
	};
	for (const [input, property] of [
		[gap, 'barGapWidth'],
		[overlap, 'barOverlap'],
	] as const) {
		input.addEventListener('change', () => {
			if (!ctx.commands.isEnabled('chart.format-series')) return refresh();
			const value = input.valueAsNumber;
			if (!Number.isInteger(value) || !input.checkValidity()) {
				ctx.toast(
					ctx.t('Enter a whole number from {min} to {max}.', { min: input.min, max: input.max }),
					'warning',
				);
				refresh();
				return;
			}
			if (activeChart(ctx)?.chart[property] !== value)
				editChart(ctx, () => ({ [property]: value }));
			refresh();
		});
	}
	relocalize();
	ctx.dialogs.register('format-chart-series', async () => {
		if (!ctx.commands.isEnabled('chart.format-series')) return;
		element.hidden = false;
		refresh();
		overlap.focus();
	});
	return { element, refresh, relocalize };
}
