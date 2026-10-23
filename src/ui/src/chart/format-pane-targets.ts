/**
 * Captions of the chart elements a Format pane can target, in the words
 * PowerPoint and Excel use in the element picker and the pane heading. Moved
 * from PowerPoint's `render/chart-format/chart-format-targets.ts`; the
 * chart-specific facts (axis orientation, series names) come from the source.
 *
 * @module chart/format-pane-targets
 */
import type { ChartFormatTarget } from 'ooxml-core/chart';
import { chartFormatTargetKey } from 'ooxml-core/chart';

import { t } from './format-pane-common';
import type { ChartFormatTargetOption, ChartFormatText } from './format-pane-types';
import type { ChartFormatAxisInfo } from './format-pane-view';

/** The chart facts captions need. */
export interface ChartFormatDescribe {
	axis(axisIndex: number): ChartFormatAxisInfo | undefined;
	seriesName(seriesIndex: number): string;
}

const text = (
	key: string,
	fallback: string,
	params?: Readonly<Record<string, string | number>>,
): ChartFormatText => (params ? { ...t(key, fallback), params } : t(key, fallback));

/** "Vertical (Value) Axis", "Secondary Horizontal (Category) Axis", ... */
function axisCaption(info: ChartFormatAxisInfo | undefined): ChartFormatText {
	if (!info) return text('target.axis', 'Axis');
	if (info.kind === 'ser') return text('target.depthAxis', 'Depth (Series) Axis');
	const value = info.kind === 'val';
	const orientation = info.vertical ? 'Vertical' : 'Horizontal';
	const id = info.secondary ? `secondary${orientation}` : orientation.toLowerCase();
	const english = `${info.secondary ? 'Secondary ' : ''}${orientation} (${value ? 'Value' : 'Category'}) Axis`;
	return text(`target.axis.${id}${value ? 'Value' : 'Category'}`, english);
}

/** The picker caption of `target`. */
export function chartFormatTargetLabel(
	target: ChartFormatTarget,
	describe: ChartFormatDescribe,
): ChartFormatText {
	switch (target.kind) {
		case 'plotArea':
			return text('target.plotArea', 'Plot Area');
		case 'title':
			return text('target.title', 'Chart Title');
		case 'legend':
			return text('target.legend', 'Legend');
		case 'axis':
			return axisCaption(describe.axis(target.axisIndex));
		case 'majorGridlines':
		case 'minorGridlines': {
			const major = target.kind === 'majorGridlines';
			return {
				...text(
					major ? 'target.majorGridlines' : 'target.minorGridlines',
					major ? '{{axis}} Major Gridlines' : '{{axis}} Minor Gridlines',
				),
				paramTexts: { axis: axisCaption(describe.axis(target.axisIndex)) },
			};
		}
		case 'series':
			return text('target.series', 'Series "{{name}}"', {
				name: describe.seriesName(target.seriesIndex),
			});
		case 'dataLabels':
			return target.seriesIndex === undefined
				? text('target.dataLabels', 'Data Labels')
				: text('target.seriesDataLabels', 'Series "{{name}}" Data Labels', {
						name: describe.seriesName(target.seriesIndex),
					});
		default:
			return text('target.chartArea', 'Chart Area');
	}
}

/** The pane heading for `target` ("Format Chart Area", "Format Axis", ...). */
export function chartFormatPaneTitle(target: ChartFormatTarget): ChartFormatText {
	switch (target.kind) {
		case 'plotArea':
			return text('title.plotArea', 'Format Plot Area');
		case 'title':
			return text('title.title', 'Format Chart Title');
		case 'legend':
			return text('title.legend', 'Format Legend');
		case 'axis':
			return text('title.axis', 'Format Axis');
		case 'majorGridlines':
			return text('title.majorGridlines', 'Format Major Gridlines');
		case 'minorGridlines':
			return text('title.minorGridlines', 'Format Minor Gridlines');
		case 'series':
			return text('title.series', 'Format Data Series');
		case 'dataLabels':
			return text('title.dataLabels', 'Format Data Labels');
		default:
			return text('title.chartArea', 'Format Chart Area');
	}
}

/** Picker entries for `targets`, captioned. */
export function chartFormatTargetOptions(
	targets: readonly ChartFormatTarget[],
	describe: ChartFormatDescribe,
): ChartFormatTargetOption[] {
	return targets.map((target) => ({
		key: chartFormatTargetKey(target),
		target,
		label: chartFormatTargetLabel(target, describe),
	}));
}

/** The chart area entry, what a stale or missing key falls back to. */
export const CHART_AREA_TARGET: ChartFormatTargetOption = {
	key: 'chartArea',
	target: { kind: 'chartArea' },
	label: text('target.chartArea', 'Chart Area'),
};
