import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import { recolorPaletteGradient, sameDrawingColor } from '../../chart/edit/paint';
import type { ChartObject } from '../model';
import type { ChartPatch } from './charts';

/** Equality of DrawingML colors, including the ordered transform sequence. */
export const sameChartColor = sameDrawingColor;

/** Change Colors preserves custom fills, as native Excel's ChartColor property does. */
export function chartColorPalettePatch(chart: ChartObject, id: number): ChartPatch | undefined {
	const next = findChartColorPalette(id);
	const previous = findChartColorPalette(chart.colorPalette ?? 10);
	if (!next) return undefined;
	const radial = chart.chartType === 'pie' || chart.chartType === 'doughnut';
	const series = chart.series.map((series, index) => {
		const copy = structuredClone(series);
		if (!previous) return copy;
		const previousChoice = chartPaletteSeriesColorChoice(previous, index, chart.series.length);
		const nextChoice = chartPaletteSeriesColorChoice(next, index, chart.series.length);
		for (const marker of [copy.marker, ...Object.values(copy.pointMarkers ?? {})]) {
			const shape = marker?.spPr;
			if (!shape) continue;
			for (const paint of [shape, shape.line]) {
				if (paint?.fill?.kind === 'solid' && sameDrawingColor(paint.fill.color, previousChoice))
					paint.fill = { kind: 'solid', color: structuredClone(nextChoice) };
			}
		}
		const fill = recolorPaletteGradient(copy.fill, previousChoice, nextChoice);
		if (fill) copy.fill = fill;
		if (sameDrawingColor(series.drawingColor, previousChoice)) {
			copy.drawingColor = nextChoice;
			delete copy.color;
		}
		if (radial && copy.pointColors) {
			for (const [key, color] of Object.entries(copy.pointColors)) {
				const point = Number(key);
				if (
					sameDrawingColor(
						color,
						chartPaletteSeriesColorChoice(previous, point, copy.values.length),
					)
				)
					copy.pointColors[point] = chartPaletteSeriesColorChoice(next, point, copy.values.length);
			}
		}
		for (const [key, value] of Object.entries(copy.pointFills ?? {})) {
			const point = Number(key);
			copy.pointFills![point] = recolorPaletteGradient(
				value,
				radial
					? chartPaletteSeriesColorChoice(previous, point, copy.values.length)
					: previousChoice,
				radial ? chartPaletteSeriesColorChoice(next, point, copy.values.length) : nextChoice,
			)!;
		}
		return copy;
	});
	return { colorPalette: id, series };
}
