import type { DiagramColor } from '../../diagram/types';
import { withDrawingColorBrightness } from '../../drawingml/drawing-color-brightness';
import type { ChartObject, Color } from '../model';
import { THEME_SLOTS } from '../layout/colors';
import type { ChartPatch } from './charts';
import { sameChartColor } from './chart-colors';
import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../../chart/color-palettes';

/** Spreadsheet theme tints use HSL luminance, represented by DrawingML lumMod/lumOff. */
export function chartDrawingColor(color: Color | undefined): DiagramColor | undefined {
	if (!color) return undefined;
	let choice: DiagramColor;
	if (color.rgb && /^(?:[0-9a-f]{2})?[0-9a-f]{6}$/i.test(color.rgb))
		choice = { kind: 'srgb', value: color.rgb.slice(-6).toUpperCase(), transforms: [] };
	else if (color.theme !== undefined && Number.isInteger(color.theme) && THEME_SLOTS[color.theme])
		choice = { kind: 'scheme', value: THEME_SLOTS[color.theme]!, transforms: [] };
	else return undefined;
	const tint = color.tint ?? 0;
	if (!Number.isFinite(tint) || tint < -1 || tint > 1) return undefined;
	return tint !== 0 ? withDrawingColorBrightness(choice, tint * 100) : choice;
}

/** Native series fill edits replace point paints, retaining effects, references and other series. */
export function chartSeriesFillPatch(
	chart: ChartObject,
	index: number,
	color: Color | null,
): ChartPatch | undefined {
	const current = chart.series[index];
	if (!Number.isInteger(index) || !current)
		throw new RangeError(`No chart series at index ${index}`);
	const choice = color === null ? undefined : chartDrawingColor(color);
	if (color !== null && !choice) throw new RangeError('Invalid chart fill color');
	const previous = current.fill?.kind === 'solid' ? current.fill.color : current.drawingColor;
	if (choice)
		choice.transforms.push(
			...structuredClone(
				previous?.transforms.filter((transform) =>
					['alpha', 'alphaMod', 'alphaOff'].includes(transform.name),
				) ?? [],
			),
		);
	return replaceSeriesPaint(chart, index, choice);
}

/** Flatten the primary paint to solid without converting a translucent CSS string back to RGB. */
export function chartSeriesSolidFillPatch(
	chart: ChartObject,
	index: number,
): ChartPatch | undefined {
	const current = chart.series[index];
	if (!Number.isInteger(index) || !current)
		throw new RangeError(`No chart series at index ${index}`);
	if (!current.fill || current.fill.kind === 'solid') return undefined;
	const choice = current.fill.kind === 'gradient' ? current.fill.stops[0]?.color : undefined;
	return replaceSeriesPaint(
		chart,
		index,
		choice ??
			current.drawingColor ??
			chartDrawingColor(current.color) ??
			chartPaletteSeriesColorChoice(
				findChartColorPalette(chart.colorPalette ?? 10) ?? findChartColorPalette(10)!,
				index,
				chart.series.length,
			),
	);
}

function replaceSeriesPaint(
	chart: ChartObject,
	index: number,
	choice: DiagramColor | undefined,
): ChartPatch | undefined {
	const current = chart.series[index]!;
	if (
		!Object.keys(current.pointColors ?? {}).length &&
		!Object.keys(current.pointFills ?? {}).length &&
		(!choice
			? current.fill?.kind === 'none'
			: !current.fill && sameChartColor(current.drawingColor, choice))
	)
		return undefined;
	const series = structuredClone(chart.series);
	const next = series[index]!;
	delete next.fill;
	delete next.color;
	delete next.drawingColor;
	delete next.pointColors;
	delete next.pointFills;
	if (choice) next.drawingColor = structuredClone(choice);
	else next.fill = { kind: 'none' };
	return { series };
}
