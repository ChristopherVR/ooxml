import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import {
	editChartGradientFill,
	type ChartGradientFillEdit,
	type ChartGradientFillValue,
} from '../../chart/edit/gradient';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartObject, Color } from '../model';
import type { ChartPatch } from './charts';
import { chartDrawingColor } from './chart-series-fill';
import { chartSeriesSolidColor } from './chart-series-transparency';
import { chartSeriesAt, cloneChartSeriesForEdit } from './chart-series-target';

export { chartGradientStopTransparency } from '../../chart/edit/gradient';

/** A gradient edit with stop colours given as spreadsheet colours. */
export type ChartGradientEdit = ChartGradientFillEdit<Color>;

/** Native gradient edits retain geometry, source flags and unrelated chart properties. */
export function chartSeriesGradientPatch(
	chart: ChartObject,
	index: number,
	edit: ChartGradientEdit,
): { patch: ChartPatch; stopIndex: number } | undefined {
	const current = chartSeriesAt(chart, index);
	const result = chartGradientFillEdit(
		current.fill,
		edit,
		chartSeriesSolidColor(chart, index) ??
			chartPaletteSeriesColorChoice(
				findChartColorPalette(chart.colorPalette ?? 10) ?? findChartColorPalette(10)!,
				index,
				chart.series.length,
			),
	);
	if (!result) return undefined;
	const { series, target: next } = cloneChartSeriesForEdit(chart, index);
	next.fill = result.fill;
	delete next.color;
	delete next.drawingColor;
	delete next.pointColors;
	delete next.pointFills;
	return { patch: { series }, stopIndex: result.stopIndex };
}

/** Spreadsheet colour of a gradient stop as DrawingML (theme tints become luminance). */
function stopColor(color: Color): DrawingColor {
	const choice = chartDrawingColor(color);
	if (!choice) throw new RangeError('Invalid gradient color');
	return choice;
}

/** Shared chart-fill edit (`chart/edit`), with stop colours given as spreadsheet colours. */
export function chartGradientFillEdit(
	old: DrawingFill | undefined,
	edit: ChartGradientEdit,
	defaultColor: DrawingColor,
): { fill: ChartGradientFillValue; stopIndex: number } | undefined {
	return editChartGradientFill(old, edit, defaultColor, stopColor);
}
