// Chart style and colour edits. The built-in style id is `c:style` (1..48) or Office's `c14:style`
// (101..148) with its `c:style` fallback, as xlsx writes it. A colour style (Change Colors) lives in
// its own part (`colorsN.xml`), not in the chart part: `applyChartColorPalette` recolours the series
// that still show the old palette's colours, keeping custom fills, as Excel's ChartColor does (lifted
// from xlsx `edit/chart-colors.ts`); the caller records the new palette id in the colour style part.
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../color-palettes';
import type { ChartSpace } from '../model';
import type { ChartShapeProperties } from '../model-series';
import { chartSeriesEntries, editChartSpace } from './draft';
import { recolorPaletteGradient, sameDrawingColor } from './paint';

/** Sets the built-in chart style: 1..48 (`c:style`) or 101..148 (`c14:style` with its fallback). */
export function setChartStyle(space: ChartSpace, id: number): ChartSpace {
	const legacy = Number.isInteger(id) && id >= 1 && id <= 48;
	const modern = Number.isInteger(id) && id >= 101 && id <= 148;
	if (!legacy && !modern) throw new RangeError('Chart style must be 1..48 or 101..148');
	return editChartSpace(space, (draft) => {
		if (modern) {
			draft.c14Style = id;
			draft.style = id - 100;
		} else {
			delete draft.c14Style;
			draft.style = id;
		}
	});
}

function recolorFill(fill: DrawingFill | undefined, previous: DrawingColor, next: DrawingColor) {
	if (fill?.kind === 'solid')
		return sameDrawingColor(fill.color, previous)
			? { kind: 'solid' as const, color: structuredClone(next) }
			: fill;
	return recolorPaletteGradient(fill, previous, next);
}

function recolorShape(
	shape: ChartShapeProperties | undefined,
	previous: DrawingColor,
	next: DrawingColor,
): void {
	if (!shape) return;
	const fill = recolorFill(shape.fill, previous, next);
	if (fill) shape.fill = fill;
	const line = shape.line && recolorFill(shape.line.fill, previous, next);
	if (shape.line && line) shape.line.fill = line;
}

/**
 * Recolours from palette `from` to palette `to` (`chart/color-palettes`): every series, marker and
 * point paint that still is the old palette's colour for its position takes the new one; custom
 * colours stay. Pie and doughnut charts that vary colours by point recolour each point.
 */
export function applyChartColorPalette(space: ChartSpace, from: number, to: number): ChartSpace {
	const previous = findChartColorPalette(from);
	const next = findChartColorPalette(to);
	if (!next) throw new RangeError(`No chart colour palette ${to}`);
	if (!previous) return space;
	return editChartSpace(space, (draft) => {
		const entries = chartSeriesEntries(draft);
		entries.forEach(({ group, series }, index) => {
			const count = entries.length;
			const old = chartPaletteSeriesColorChoice(previous, index, count);
			const choice = chartPaletteSeriesColorChoice(next, index, count);
			recolorShape(series.spPr, old, choice);
			recolorShape(series.marker?.spPr, old, choice);
			const radial =
				(group.kind === 'pie' || group.kind === 'doughnut') && group.varyColors !== false;
			const points = series.values?.cache?.pointCount ?? series.values?.cache?.points.length ?? 0;
			for (const point of series.dataPoints) {
				const at = point.index ?? 0;
				recolorShape(
					point.spPr,
					radial ? chartPaletteSeriesColorChoice(previous, at, points) : old,
					radial ? chartPaletteSeriesColorChoice(next, at, points) : choice,
				);
				recolorShape(point.marker?.spPr, old, choice);
			}
		});
	});
}
