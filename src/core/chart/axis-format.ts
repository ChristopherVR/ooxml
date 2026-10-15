import { formatChartNumber } from './number-format';

/**
 * PowerPoint's compact axis-value formatting (1.2K / 3.4M / integer / one-decimal).
 *
 * `formatCode` is the chart's own `c:numFmt/@formatCode`. When the source declares one
 * (`0%`, `#,##0`, `$#,##0.00`) it wins: the cached values behind a percentage chart are
 * fractions, and the compact fallback renders them as `0.5` where Office shows `50%`. Codes
 * outside the supported subset fall through to the compact form.
 *
 * Distinct from the Excel painter's `formatAxisValue` (`chart/render/chart-scale.ts`), which
 * groups thousands instead of abbreviating them.
 */
export function formatCompactAxisValue(value: number, formatCode?: string): string {
	const formatted = formatChartNumber(value, formatCode);
	if (formatted !== undefined) return formatted;
	if (Math.abs(value) >= 1_000_000) return `${(value / 1_000_000).toFixed(1)}M`;
	if (Math.abs(value) >= 1_000) return `${(value / 1_000).toFixed(1)}K`;
	if (Number.isInteger(value)) return String(value);
	return value.toFixed(1);
}
