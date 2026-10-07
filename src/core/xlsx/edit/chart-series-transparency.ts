import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import { resolveDrawingColor } from '../../diagram/drawing-color';
import type { DiagramColor } from '../../diagram/types';
import type { ChartObject } from '../model';
import type { ChartPatch } from './charts';
import { chartDrawingColor } from './chart-series-fill';
import { sameChartColor } from './chart-colors';

export function chartSeriesSolidColor(chart: ChartObject, index: number): DiagramColor | undefined {
	const series = chart.series[index];
	if (!series || (series.fill && series.fill.kind !== 'solid')) return undefined;
	return series.fill?.kind === 'solid'
		? series.fill.color
		: (series.drawingColor ??
				chartDrawingColor(series.color) ??
				chartPaletteSeriesColorChoice(
					findChartColorPalette(chart.colorPalette ?? 10) ?? findChartColorPalette(10)!,
					index,
					chart.series.length,
				));
}

/** Theme lookup cannot affect opacity; the shared resolver applies ordered alpha transforms. */
export function chartSeriesTransparency(chart: ChartObject, index: number): number | undefined {
	const color = chartSeriesSolidColor(chart, index);
	if (!color) return undefined;
	const resolved = resolveDrawingColor(
		color,
		{ scheme: () => '#000000' },
		{ transformOrder: 'document' },
	);
	return resolved ? Math.round((1 - resolved.alpha) * 100000) / 1000 : undefined;
}

/** Set the solid series opacity while retaining its theme and luminance transforms. */
export function chartSeriesTransparencyPatch(
	chart: ChartObject,
	index: number,
	percent: number,
): ChartPatch | undefined {
	if (!Number.isFinite(percent) || percent < 0 || percent > 100)
		throw new RangeError('Chart transparency must be from 0 to 100');
	const current = chart.series[index];
	if (!Number.isInteger(index) || !current)
		throw new RangeError(`No chart series at index ${index}`);
	const color = chartSeriesSolidColor(chart, index);
	if (!color) return undefined;
	const choice = structuredClone(color);
	choice.transforms = choice.transforms.filter(
		(transform) => !['alpha', 'alphaMod', 'alphaOff'].includes(transform.name),
	);
	if (percent !== 0)
		choice.transforms.push({ name: 'alpha', value: String(Math.round((100 - percent) * 1000)) });
	if (
		sameChartColor(choice, color) &&
		!Object.keys(current.pointColors ?? {}).length &&
		!Object.keys(current.pointFills ?? {}).length
	)
		return undefined;
	const series = structuredClone(chart.series);
	const next = series[index]!;
	delete next.fill;
	delete next.color;
	delete next.pointColors;
	delete next.pointFills;
	next.drawingColor = choice;
	return { series };
}
