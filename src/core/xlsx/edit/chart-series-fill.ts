import type { DiagramColor } from '../../diagram/types';
import type { ChartObject, Color } from '../model';
import { THEME_SLOTS } from '../layout/colors';
import type { ChartPatch } from './charts';
import { sameChartColor } from './chart-colors';

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
	if (tint !== 0) {
		choice.transforms.push({
			name: 'lumMod',
			value: String(Math.round((1 - Math.abs(tint)) * 100000)),
		});
		if (tint > 0)
			choice.transforms.push({ name: 'lumOff', value: String(Math.round(tint * 100000)) });
	}
	return choice;
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
	if (
		!Object.keys(current.pointColors ?? {}).length &&
		!Object.keys(current.pointFills ?? {}).length &&
		(color === null
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
	if (choice) next.drawingColor = choice;
	else next.fill = { kind: 'none' };
	return { series };
}
