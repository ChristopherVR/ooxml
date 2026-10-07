import { fit, textWidth, type text } from './chart-svg-util';

export const CHART_FONT_FAMILY = 'Calibri, Carlito, Arial, sans-serif';
type TextAttributes = NonNullable<Parameters<typeof text>[3]>;

export interface ChartSvgOptions {
	/** Host width in CSS pixels for the painter's exact CSS font. Core never accesses a DOM. */
	measureText?: (content: string, font: string) => number;
}

/** Use the same family, size and emphasis for measurement and SVG painting. */
export function chartTextWidth(
	content: string,
	attrs: TextAttributes,
	options: ChartSvgOptions,
): number {
	const font = `${attrs.italic ? 'italic' : 'normal'} ${attrs.weight ?? 'normal'} ${attrs.size ?? 10}px ${attrs.family ?? CHART_FONT_FAMILY}`;
	const width = options.measureText?.(content, font);
	return width !== undefined && Number.isFinite(width) && width >= 0
		? width
		: textWidth(content, attrs.size);
}

/** Fit whole Unicode code points with host metrics when available. */
export function fitChartText(
	content: string,
	maxWidth: number,
	attrs: TextAttributes,
	options: ChartSvgOptions,
): string {
	if (!options.measureText) return fit(content, maxWidth, attrs.size);
	if (chartTextWidth(content, attrs, options) <= maxWidth) return content;
	const points = [...content];
	let low = 0;
	let high = points.length;
	while (low < high) {
		const mid = Math.ceil((low + high) / 2);
		if (chartTextWidth(`${points.slice(0, mid).join('')}…`, attrs, options) <= maxWidth) low = mid;
		else high = mid - 1;
	}
	return `${points.slice(0, low).join('')}…`;
}
