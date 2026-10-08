import { fit, textWidth, type text } from './chart-svg-util';

export const CHART_FONT_FAMILY = 'Calibri, Carlito, Arial, sans-serif';
type TextAttributes = NonNullable<Parameters<typeof text>[3]>;

export interface ChartSvgOptions {
	/** Host width in CSS pixels for the painter's exact CSS font. Core never accesses a DOM. */
	measureText?: (content: string, font: string) => number;
	/** Natural font ascent/descent in CSS pixels, when the host has vertical metrics. */
	measureFont?: (font: string) => { ascent: number; descent: number } | undefined;
}

function cssFont(attrs: TextAttributes): string {
	return `${attrs.italic ? 'italic' : 'normal'} ${attrs.weight ?? 'normal'} ${attrs.size ?? 10}px ${attrs.family ?? CHART_FONT_FAMILY}`;
}

export function chartFontMetrics(
	attrs: TextAttributes,
	options: ChartSvgOptions,
): { ascent: number; descent: number } | undefined {
	const metrics = options.measureFont?.(cssFont(attrs));
	return metrics &&
		Number.isFinite(metrics.ascent) &&
		Number.isFinite(metrics.descent) &&
		metrics.ascent >= 0 &&
		metrics.descent >= 0 &&
		metrics.ascent + metrics.descent > 0
		? metrics
		: undefined;
}

/** Use the same family, size and emphasis for measurement and SVG painting. */
export function chartTextWidth(
	content: string,
	attrs: TextAttributes,
	options: ChartSvgOptions,
): number {
	const width = options.measureText?.(content, cssFont(attrs));
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
