import {
	chartSeriesFillPatch,
	chartSeriesSolidFillPatch,
	type ChartObject,
	type ChartViewModel,
	type Color,
} from 'ooxml-core/xlsx';
import { activeChart, type EditorContext } from 'ooxml-core/xlsx/ui';
import { el, field, select } from './dialogs/fields';
import { openColorGrid } from './ribbon/color-grid';
import { createSeriesTransparency } from './chart-series-transparency';

/** Primary series fill controls reuse the ribbon's themed color picker and core paint edits. */
export function createSeriesFill(ctx: EditorContext, selected: () => number) {
	const element = el(ctx, 'div', 'xve-chart-series-fill');
	const heading = el(ctx, 'h3');
	const kind = select(ctx, []);
	const kindField = field(ctx, 'Fill', kind);
	const color = el(ctx, 'button', 'xve-input xve-chart-series-color');
	color.type = 'button';
	const colorField = field(ctx, 'Color', color);
	const transparency = createSeriesTransparency(ctx, selected);
	element.append(heading, kindField, colorField, transparency.element);
	let current: ChartObject | undefined;
	const apply = (value: Color | null) => {
		if (!current || !ctx.commands.isEnabled('chart.format-series')) return;
		const found = activeChart(ctx);
		if (!found || found.chart !== current) return;
		const patch = chartSeriesFillPatch(current, selected(), value);
		if (patch) ctx.session()?.updateChart(ctx.activeSheet(), found.index, patch);
	};
	kind.addEventListener('change', () => {
		if (kind.value === 'none') apply(null);
		else if (kind.value === 'solid') {
			if (!current || !ctx.commands.isEnabled('chart.format-series')) return;
			const found = activeChart(ctx);
			if (!found || found.chart !== current) return;
			const patch = chartSeriesSolidFillPatch(current, selected());
			if (patch) ctx.session()?.updateChart(ctx.activeSheet(), found.index, patch);
		}
	});
	color.addEventListener('click', () => {
		if (!current || !ctx.commands.isEnabled('chart.format-series')) return;
		const chart = current;
		const index = selected();
		const workbook = ctx.workbook();
		openColorGrid(
			color,
			(choice) => {
				if (current === chart && ctx.workbook() === workbook && selected() === index && choice)
					apply(choice);
			},
			{ t: ctx.t, ...(workbook ? { theme: workbook.theme } : {}) },
		);
	});
	const refresh = (chart: ChartObject | undefined, model: ChartViewModel | undefined) => {
		current = chart;
		transparency.refresh(chart);
		const series = chart?.series[selected()];
		const fill = series?.fill;
		const value =
			fill?.kind === 'none'
				? 'none'
				: fill?.kind === 'gradient'
					? 'gradient'
					: fill && fill.kind !== 'solid'
						? 'imported'
						: 'solid';
		const entries = [
			['none', 'No fill'],
			['solid', 'Solid fill'],
		];
		if (value === 'gradient') entries.push(['gradient', 'Gradient fill']);
		else if (value === 'imported') entries.push(['imported', 'Imported fill']);
		kind.replaceChildren(
			...entries.map(([key, label]) => {
				const option = el(ctx, 'option');
				option.value = key!;
				option.textContent = ctx.t(label!);
				option.disabled = key === 'gradient' || key === 'imported';
				return option;
			}),
		);
		kind.value = value;
		kind.disabled = !series || !ctx.commands.isEnabled('chart.format-series');
		color.disabled = kind.disabled || value !== 'solid';
		const paint = model?.series[selected()]?.color;
		color.style.setProperty('--series-fill', paint && paint !== 'none' ? paint : 'transparent');
		heading.textContent = ctx.t('Fill');
		for (const [row, control, label] of [
			[kindField, kind, 'Fill'],
			[colorField, color, 'Color'],
		] as const) {
			row.querySelector('span')!.textContent = ctx.t(label);
			control.setAttribute('aria-label', ctx.t(label));
		}
	};
	return { element, refresh };
}
