import {
	chartSeriesFillPatch,
	chartSeriesSolidFillPatch,
	chartSeriesGradientPatch,
	chartSeriesTransparency,
	chartSeriesTransparencyPatch,
	chartElementFill,
	chartElementFillPatch,
	chartElementSolidFillPatch,
	chartElementGradientPatch,
	chartElementTransparency,
	chartElementTransparencyPatch,
	chartWithElementFills,
	type ChartObject,
	type ChartViewModel,
	type ChartPatch,
	type Color,
	type ChartGradientEdit,
	type ChartFillPart,
} from 'ooxml-core/xlsx';
import type { DrawingFill } from 'ooxml-core/drawingml';
import type { ChartGradientFill } from 'ooxml-core/chart';

/** UI target binding; every model edit and temporary model remains in core. */
export interface ChartFillBinding {
	commandId: string;
	key(): number | ChartFillPart;
	fill(chart: ChartObject): DrawingFill | undefined;
	gradient(view: ChartViewModel): ChartGradientFill | undefined;
	paint(view: ChartViewModel): string | undefined;
	rectangular(chart: ChartObject): boolean;
	color(chart: ChartObject, color: Color | null): ChartPatch | undefined;
	solid(chart: ChartObject): ChartPatch | undefined;
	edit(
		chart: ChartObject,
		edit: ChartGradientEdit,
	): { patch: ChartPatch; stopIndex: number } | undefined;
	transparency(chart: ChartObject): number | undefined;
	setTransparency(chart: ChartObject, value: number): ChartPatch | undefined;
	preview(chart: ChartObject, patch: ChartPatch): ChartObject;
}

export function seriesFillBinding(selected: () => number): ChartFillBinding {
	return {
		commandId: 'chart.format-series',
		key: selected,
		fill: (chart) => chart.series[selected()]?.fill,
		gradient: (view) => view.series[selected()]?.gradient,
		paint: (view) => view.series[selected()]?.color,
		rectangular: (chart) => chart.chartType === 'bar' || chart.chartType === 'column',
		color: (chart, color) => chartSeriesFillPatch(chart, selected(), color),
		solid: (chart) => chartSeriesSolidFillPatch(chart, selected()),
		edit: (chart, edit) => chartSeriesGradientPatch(chart, selected(), edit),
		transparency: (chart) => chartSeriesTransparency(chart, selected()),
		setTransparency: (chart, value) => chartSeriesTransparencyPatch(chart, selected(), value),
		preview: (chart, patch) => ({ ...chart, ...patch }),
	};
}

export function backgroundFillBinding(selected: () => ChartFillPart): ChartFillBinding {
	return {
		commandId: 'chart.format-area',
		key: selected,
		fill: (chart) => chartElementFill(chart, selected()),
		gradient: (view) => view.appearance?.[selected()]?.gradient,
		paint: (view) =>
			view.appearance?.[selected()]?.fillColor ?? (selected() === 'chartArea' ? '#FFFFFF' : 'none'),
		rectangular: () => true,
		color: (chart, color) => chartElementFillPatch(chart, selected(), color),
		solid: (chart) => chartElementSolidFillPatch(chart, selected()),
		edit: (chart, edit) => chartElementGradientPatch(chart, selected(), edit),
		transparency: (chart) => chartElementTransparency(chart, selected()),
		setTransparency: (chart, value) => chartElementTransparencyPatch(chart, selected(), value),
		preview: (chart, patch) => chartWithElementFills(chart, patch.elementFills ?? {}),
	};
}
