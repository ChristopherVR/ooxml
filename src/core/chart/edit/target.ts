// The chart element a Format pane formats, and its stable string key (`chartArea`, `axis:1`,
// `series:0`, `dataLabels`, `dataLabels:2`...). Moved from PowerPoint's
// `utils/chart-element-format-keys.ts` so every product's pane addresses elements the same way.
// Axis and gridline targets name the axis by its position in the axis list (`plotArea.axes` of a
// `ChartSpace`); series and data-label targets by the flattened series index (see `draft.ts`). A
// data-label target without a series covers every series (the chart-level "Data Labels").

export type ChartFormatTarget =
	| { kind: 'chartArea' }
	| { kind: 'plotArea' }
	| { kind: 'title' }
	| { kind: 'legend' }
	| { kind: 'axis'; axisIndex: number }
	| { kind: 'majorGridlines'; axisIndex: number }
	| { kind: 'minorGridlines'; axisIndex: number }
	| { kind: 'series'; seriesIndex: number }
	| { kind: 'dataLabels'; seriesIndex?: number };

/** A {@link ChartFormatTarget} as a stable string. */
export type ChartFormatTargetKey = string;

/** The key of `target`. */
export function chartFormatTargetKey(target: ChartFormatTarget): ChartFormatTargetKey {
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

/** Parses a key back into a target, or `undefined` for an unknown key. */
export function parseChartFormatTargetKey(
	key: ChartFormatTargetKey,
): ChartFormatTarget | undefined {
	const [kind, raw] = key.split(':');
	const index = raw === undefined ? undefined : Number(raw);
	if (index !== undefined && (!Number.isInteger(index) || index < 0)) return undefined;
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
