import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartObject } from '../model';
import type { ChartPatch } from './charts';

/** Equality of DrawingML colors, including the ordered transform sequence. */
export function sameChartColor(a: DrawingColor | undefined, b: DrawingColor | undefined): boolean {
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

/** Palette transforms precede gradient shading transforms in Excel's saved XML. */
function recolorGradient(
	fill: DrawingFill | undefined,
	previous: DrawingColor,
	next: DrawingColor,
): DrawingFill | undefined {
	if (fill?.kind !== 'gradient') return fill;
	const matches = fill.stops.every((stop) =>
		sameChartColor(
			{ ...stop.color, transforms: stop.color.transforms.slice(0, previous.transforms.length) },
			previous,
		),
	);
	if (!matches) return fill;
	return {
		...fill,
		stops: fill.stops.map((stop) => ({
			...stop,
			color: {
				...next,
				transforms: [
					...next.transforms,
					...stop.color.transforms.slice(previous.transforms.length),
				],
			},
		})),
	};
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
		const previousChoice = chartPaletteSeriesColorChoice(previous, index, chart.series.length);
		const nextChoice = chartPaletteSeriesColorChoice(next, index, chart.series.length);
		const fill = recolorGradient(copy.fill, previousChoice, nextChoice);
		if (fill) copy.fill = fill;
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
		for (const [key, value] of Object.entries(copy.pointFills ?? {})) {
			const point = Number(key);
			copy.pointFills![point] = recolorGradient(
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
