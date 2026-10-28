import { chartDataTableMetrics } from '../data-table-layout';
import { chartPointsToPixels } from './chart-appearance';
import { chartTextAttributes } from './chart-svg-appearance';
import { strokeAttributes } from './chart-svg-analysis';
import { formatChartNumber } from '../number-format';
import { formatAxisValue } from './chart-scale';
import { text, rect, fit, n, type Rect } from './chart-svg-util';
import type { ChartViewModel } from './chart-view';

const rows = (model: ChartViewModel) => [
	...model.series,
	...(model.extraGroups ?? []).flatMap((g) => g.series),
];
export function chartTableMetrics(model: ChartViewModel) {
	if (!model.dataTable) return undefined;
	const size = model.dataTable.text?.fontSize;
	return chartDataTableMetrics(
		size === undefined ? 13.333333333333334 : chartPointsToPixels(size),
		rows(model).map((s) => s.name),
		model.dataTable.source.showKeys !== false,
	);
}

/** Table columns occupy the same category bands as the plot, with keys hanging to its left. */
export function chartDataTableSvg(model: ChartViewModel, plot: Rect): string {
	const table = model.dataTable;
	const metrics = chartTableMetrics(model);
	if (!table || !metrics) return '';
	const { source } = table;
	const { rowH, keyW, height, fontSize } = metrics;
	const series = rows(model);
	const top = plot.y + plot.h;
	const left = plot.x - keyW;
	const right = plot.x + plot.w;
	const bottom = top + height;
	const cellW = plot.w / Math.max(1, model.categories.length);
	const styled = { ...model, appearance: { ...model.appearance, dataLabel: table.text ?? {} } };
	const attrs = chartTextAttributes(styled, 'dataLabel', fontSize);
	const out = ['<g data-chart-part="dataTable">'];
	const border = (x1: number, y1: number, x2: number, y2: number) =>
		out.push(
			`<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" ${strokeAttributes(table.stroke)}/>`,
		);
	const label = (
		x: number,
		y: number,
		value: string,
		width: number,
		anchor: 'middle' | 'start' = 'middle',
	) =>
		out.push(
			text(x, y + fontSize * 0.35, fit(value, Math.max(1, width), fontSize), { ...attrs, anchor }),
		);
	if (table.fill) out.push(rect(left, top, right - left, height, table.fill));
	if (source.showOutline !== false) {
		border(plot.x, top, right, top);
		border(right, top, right, bottom);
		border(right, bottom, left, bottom);
		border(left, bottom, left, top + rowH);
		border(left, top + rowH, plot.x, top + rowH);
		border(plot.x, top, plot.x, top + rowH);
	}
	model.categories.forEach((category, i) => {
		label(plot.x + (i + 0.5) * cellW, top + rowH / 2, category, cellW - 4);
		if (source.showVerticalBorder !== false && i > 0)
			border(plot.x + i * cellW, top, plot.x + i * cellW, bottom);
	});
	if (source.showVerticalBorder !== false) border(plot.x, top, plot.x, bottom);
	series.forEach((s, i) => {
		const rowTop = top + rowH * (i + 1);
		const cy = rowTop + rowH / 2;
		if (source.showHorizontalBorder !== false) border(left, rowTop, right, rowTop);
		if (source.showKeys !== false) {
			const swatch = Math.max(6, Math.round(fontSize * 0.55));
			out.push(rect(left + 4, cy - swatch / 2, swatch, swatch, s.color));
		}
		const keyInset = source.showKeys !== false ? Math.max(6, Math.round(fontSize * 0.55)) + 7 : 4;
		label(left + keyInset, cy, s.name, keyW - keyInset - 4, 'start');
		model.categories.forEach((_, j) => {
			const value = s.values[j];
			const formatted =
				value === null || value === undefined
					? ''
					: (formatChartNumber(value, s.numberFormat) ?? formatAxisValue(value));
			label(plot.x + (j + 0.5) * cellW, cy, formatted, cellW - 4);
		});
	});
	return out.join('') + '</g>';
}
