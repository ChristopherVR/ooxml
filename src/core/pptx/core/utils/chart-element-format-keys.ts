/**
 * String keys for chart Format-pane targets (`chartArea`, `axis:1`,
 * `series:0`, `dataLabels`, ...), the keys of `PptxChartData.elementFormats`.
 *
 * @module utils/chart-element-format-keys
 */
import type { PptxChartFormatTarget, PptxChartFormatTargetKey } from '../types';

/** The record key of `target`. */
export function chartFormatTargetKey(target: PptxChartFormatTarget): PptxChartFormatTargetKey {
	switch (target.kind) {
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines':
			return `${target.kind}:${target.axisIndex}`;
		case 'series':
			return `series:${target.seriesIndex}`;
		case 'dataLabels':
			return target.seriesIndex === undefined ? 'dataLabels' : `dataLabels:${target.seriesIndex}`;
		default:
			return target.kind;
	}
}

/** Parse a record key back into a target, or `undefined` for an unknown key. */
export function parseChartFormatTargetKey(
	key: PptxChartFormatTargetKey,
): PptxChartFormatTarget | undefined {
	const [kind, raw] = key.split(':');
	const index = raw === undefined ? undefined : Number(raw);
	if (index !== undefined && (!Number.isInteger(index) || index < 0)) {
		return undefined;
	}
	switch (kind) {
		case 'chartArea':
		case 'plotArea':
		case 'title':
		case 'legend':
			return index === undefined ? { kind } : undefined;
		case 'axis':
		case 'majorGridlines':
		case 'minorGridlines':
			return index === undefined ? undefined : { kind, axisIndex: index };
		case 'series':
			return index === undefined ? undefined : { kind, seriesIndex: index };
		case 'dataLabels':
			return index === undefined ? { kind } : { kind, seriesIndex: index };
		default:
			return undefined;
	}
}
