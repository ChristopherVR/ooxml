import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import type { DiagramColor } from '../../diagram/types';
import type { ChartObject } from '../model';
import type { ChartPatch } from './charts';

/** Equality of DrawingML colors, including the ordered transform sequence. */
export function sameChartColor(a: DiagramColor | undefined, b: DiagramColor | undefined): boolean {
	if (!a || !b) return a === b;
	return (
		a.kind === b.kind &&
		a.value === b.value &&
		a.fallback === b.fallback &&
		a.transforms.length === b.transforms.length &&
		a.transforms.every(
			(t, i) => b.transforms[i]?.name === t.name && b.transforms[i]?.value === t.value,
		)
	);
}

/** Change Colors preserves custom fills, as native Excel's ChartColor property does. */
export function chartColorPalettePatch(chart: ChartObject, id: number): ChartPatch | undefined {
	const next = findChartColorPalette(id);
	const previous = findChartColorPalette(chart.colorPalette ?? 10);
	if (!next) return undefined;
	const radial = chart.chartType === 'pie' || chart.chartType === 'doughnut';
	const series = chart.series.map((series, index) => {
		const copy = structuredClone(series);
		if (!previous) return copy;
		if (
			sameChartColor(
				series.drawingColor,
				chartPaletteSeriesColorChoice(previous, index, chart.series.length),
			)
		) {
			copy.drawingColor = chartPaletteSeriesColorChoice(next, index, chart.series.length);
			delete copy.color;
		}
		if (radial && copy.pointColors) {
			for (const [key, color] of Object.entries(copy.pointColors)) {
				const point = Number(key);
				if (
					sameChartColor(color, chartPaletteSeriesColorChoice(previous, point, copy.values.length))
				)
					copy.pointColors[point] = chartPaletteSeriesColorChoice(next, point, copy.values.length);
			}
		}
		return copy;
	});
	return { colorPalette: id, series };
}
