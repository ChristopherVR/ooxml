import type { ChartStylePart } from '../style-definition';
import { chartPointsToPixels } from './chart-appearance';
import type { ChartViewModel } from './chart-view';
import { AXIS_COLOR, GRID_COLOR, FONT_SIZE, n, esc, type Rect, type text } from './chart-svg-util';
import { buildChartGradientDef, type ChartGradientFill } from '../gradient-definition';
import { chartGradientMarkup } from '../gradient-markup';
import { cssFontFamily } from '../../text/css-font-family';
import { svgDropShadowElement, type DrawingSvgShadow } from '../../drawingml/drawing-shadow';

/** Each standalone SVG needs distinct gradient targets when inserted beside other charts. */
let nextPaintId = 0;
export function chartGradientPaint(
	model: ChartViewModel,
	width = 1024,
	height = 768,
): { model: ChartViewModel; defs: string } {
	const prefix = `xlsx-chart-${++nextPaintId}`;
	const appearance = { ...model.appearance };
	const defs: string[] = [];
	const shadowPaint = (shadow: DrawingSvgShadow, suffix: string) => {
		const id = `${prefix}-${suffix}`;
		const pad = Math.max(4, shadow.blur * 3 + Math.abs(shadow.dx) + Math.abs(shadow.dy));
		defs.push(
			`<filter id="${id}" filterUnits="userSpaceOnUse" x="${n(-pad)}" y="${n(-pad)}" width="${n(width + pad * 2)}" height="${n(height + pad * 2)}" color-interpolation-filters="sRGB">${svgDropShadowElement(shadow, { number: n, color: (value) => esc(value) })}</filter>`,
		);
		return `url(#${id})`;
	};
	for (const [part, entry] of Object.entries(appearance)) {
		if (entry?.textShadow)
			appearance[part as ChartStylePart] = {
				...entry,
				textShadowFilter: shadowPaint(entry.textShadow, `${part}-text-shadow`),
			};
	}
	const paint = (
		gradient: ChartGradientFill,
		suffix: string,
		bounds?: { width: number; height: number; shape: 'rect' },
	) => {
		const def = buildChartGradientDef(`${prefix}-${suffix}`, gradient, bounds);
		defs.push(chartGradientMarkup(def));
		return `url(#${def.id})`;
	};
	for (const part of ['chartArea', 'plotArea'] as const) {
		const entry = appearance[part];
		if (!entry?.gradient) continue;
		if (part === 'plotArea' && ['circle', 'shape'].includes(entry.gradient.path ?? '')) continue;
		appearance[part] = {
			...entry,
			fillColor: paint(entry.gradient, part, { width, height, shape: 'rect' }),
		};
	}
	const series = model.series.map((source, index) => {
		const view = { ...source };
		if (source.shadow) view.shadowFilter = shadowPaint(source.shadow, `s${index}-shadow`);
		const bounds =
			model.type === 'bar' || model.type === 'column'
				? { width: 1, height: 1, shape: 'rect' as const }
				: undefined;
		if (source.gradient) view.color = paint(source.gradient, `s${index}`, bounds);
		if (source.pointGradients) {
			view.pointColors = [...(source.pointColors ?? [])];
			for (const [key, gradient] of Object.entries(source.pointGradients))
				view.pointColors[Number(key)] = paint(gradient, `s${index}-p${key}`, bounds);
		}
		return view;
	});
	return defs.length
		? { model: { ...model, appearance, series }, defs: `<defs>${defs.join('')}</defs>` }
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
		...(entry?.underline === undefined ? {} : { underline: entry.underline }),
		...(entry?.typeface === undefined ? {} : { family: cssFontFamily(entry.typeface) }),
		...(entry?.textShadowFilter === undefined ? {} : { filter: entry.textShadowFilter }),
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
	part: 'chartArea' | 'plotArea' | 'title' | 'legend',
	area: Rect,
): string {
	if (area.w <= 0 || area.h <= 0) return '';
	const entry = model.appearance?.[part];
	let fill = entry?.fillColor ?? (part === 'chartArea' ? '#FFFFFF' : 'none');
	let defs = '';
	if (
		part !== 'chartArea' &&
		entry?.gradient &&
		(part !== 'plotArea' || ['circle', 'shape'].includes(entry.gradient.path ?? ''))
	) {
		const def = buildChartGradientDef(`xlsx-chart-${++nextPaintId}-${part}`, entry.gradient, {
			width: area.w,
			height: area.h,
			shape: 'rect',
		});
		defs = `<defs>${chartGradientMarkup(def)}</defs>`;
		fill = `url(#${def.id})`;
	}
	const stroke = entry?.lineColor ?? 'none';
	if (fill === 'none' && stroke === 'none') return '';
	const width = entry?.lineWidth === undefined ? 1 : chartPointsToPixels(entry.lineWidth);
	return `${defs}<rect x="${n(area.x)}" y="${n(area.y)}" width="${n(area.w)}" height="${n(area.h)}" fill="${esc(fill)}" stroke="${esc(stroke)}" stroke-width="${n(width)}" data-chart-part="${part}"/>`;
}
