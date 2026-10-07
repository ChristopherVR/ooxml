import type { ChartStylePart } from '../../chart/style-definition';
import { chartPointsToPixels } from './chart-appearance';
import type { ChartViewModel } from './chart-view';
import { AXIS_COLOR, GRID_COLOR, FONT_SIZE, n, esc, type Rect, type text } from './chart-svg-util';
import { buildChartGradientDef } from '../../chart/gradient-definition';
import { cssFontFamily } from './font-family';

/** Each standalone SVG needs distinct gradient targets when inserted beside other charts. */
let nextPaintId = 0;
export function chartGradientPaint(model: ChartViewModel): { model: ChartViewModel; defs: string } {
	const prefix = `xlsx-chart-${++nextPaintId}`;
	const appearance = { ...model.appearance };
	const defs: string[] = [];
	for (const part of ['chartArea', 'plotArea'] as const) {
		const entry = appearance[part];
		if (!entry?.gradient) continue;
		const def = buildChartGradientDef(`${prefix}-${part}`, entry.gradient);
		const geometry =
			def.kind === 'linearGradient'
				? `x1="${def.x1}" y1="${def.y1}" x2="${def.x2}" y2="${def.y2}"`
				: `cx="${def.cx}" cy="${def.cy}" r="${def.r}"`;
		const stops = def.stops
			.map(
				(stop) =>
					`<stop offset="${stop.offset}" stop-color="${esc(stop.color)}" stop-opacity="${stop.opacity ?? 1}"/>`,
			)
			.join('');
		defs.push(`<${def.kind} id="${def.id}" ${geometry}>${stops}</${def.kind}>`);
		appearance[part] = { ...entry, fillColor: `url(#${def.id})` };
	}
	return defs.length
		? { model: { ...model, appearance }, defs: `<defs>${defs.join('')}</defs>` }
		: { model, defs: '' };
}

/** Convert native points once at the CSS-pixel painter boundary. */
export function chartTextAttributes(
	model: ChartViewModel,
	part: ChartStylePart,
	fallbackSize = FONT_SIZE,
): NonNullable<Parameters<typeof text>[3]> {
	const entry = model.appearance?.[part];
	return {
		size: entry?.fontSize === undefined ? fallbackSize : chartPointsToPixels(entry.fontSize),
		...(entry?.color === undefined ? {} : { fill: entry.color }),
		...(entry?.bold === undefined ? {} : { weight: entry.bold ? 'bold' : 'normal' }),
		...(entry?.italic === undefined ? {} : { italic: entry.italic }),
		...(entry?.typeface === undefined ? {} : { family: cssFontFamily(entry.typeface) }),
	};
}

export function chartStroke(
	model: ChartViewModel,
	part: ChartStylePart,
): { color: string; width: number } {
	const entry = model.appearance?.[part];
	return {
		color:
			entry?.axisVisible === false
				? 'none'
				: (entry?.lineColor ?? (part === 'gridlineMajor' ? GRID_COLOR : AXIS_COLOR)),
		width: entry?.lineWidth === undefined ? 1 : chartPointsToPixels(entry.lineWidth),
	};
}

export function chartAreaRect(
	model: ChartViewModel,
	part: 'chartArea' | 'plotArea',
	area: Rect,
): string {
	const entry = model.appearance?.[part];
	const fill = entry?.fillColor ?? (part === 'chartArea' ? '#FFFFFF' : 'none');
	const stroke = entry?.lineColor ?? 'none';
	if (fill === 'none' && stroke === 'none') return '';
	const width = entry?.lineWidth === undefined ? 1 : chartPointsToPixels(entry.lineWidth);
	return `<rect x="${n(area.x)}" y="${n(area.y)}" width="${n(area.w)}" height="${n(area.h)}" fill="${esc(fill)}" stroke="${esc(stroke)}" stroke-width="${n(width)}"/>`;
}
