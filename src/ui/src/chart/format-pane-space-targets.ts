/**
 * The elements of a `ChartSpace` a Format pane (and a Format tab's Chart Elements list) can
 * target, captioned, in PowerPoint's picker order. Split from `format-pane-space` so a ribbon can
 * list them without loading the pane's views.
 *
 * @module chart/format-pane-space-targets
 */
import type { ChartAxis, ChartFormatTarget, ChartSpace, ChartSpaceSeries } from 'ooxml-core/chart';
import { chartSeriesEntries } from 'ooxml-core/chart';

import { chartFormatTargetOptions, type ChartFormatDescribe } from './format-pane-targets';
import type { ChartFormatTargetOption } from './format-pane-types';

const vertical = (axis: ChartAxis, space: ChartSpace): boolean => {
	if (axis.position) return axis.position === 'l' || axis.position === 'r';
	const horizontalBars = space.plotArea.groups.some((group) => group.barDirection === 'bar');
	return axis.kind === 'val' ? !horizontalBars : horizontalBars;
};

const seriesName = (series: ChartSpaceSeries | undefined, index: number) =>
	series?.tx?.text || series?.tx?.value || `Series ${index + 1}`;

/** Every element of `space` a pane can target. */
export function chartSpaceTargets(space: ChartSpace): ChartFormatTarget[] {
	const targets: ChartFormatTarget[] = [{ kind: 'chartArea' }];
	if (space.title) targets.push({ kind: 'title' });
	space.plotArea.axes.forEach((axis, axisIndex) => {
		if (axis.deleted) return;
		targets.push({ kind: 'axis', axisIndex });
		if (axis.majorGridlines) targets.push({ kind: 'majorGridlines', axisIndex });
		if (axis.minorGridlines) targets.push({ kind: 'minorGridlines', axisIndex });
	});
	if (space.legend) targets.push({ kind: 'legend' });
	targets.push({ kind: 'plotArea' });
	if (space.plotArea.groups.some((group) => group.dataLabels && !group.dataLabels.deleted))
		targets.push({ kind: 'dataLabels' });
	chartSeriesEntries(space).forEach((_, seriesIndex) => {
		targets.push({ kind: 'series', seriesIndex }, { kind: 'dataLabels', seriesIndex });
	});
	return targets;
}

/** The chart facts the captions need: axis orientation and series names. */
export function chartSpaceDescribe(space: ChartSpace): ChartFormatDescribe {
	const entries = chartSeriesEntries(space);
	return {
		axis: (axisIndex) => {
			const axis = space.plotArea.axes[axisIndex];
			if (!axis) return undefined;
			const secondary = space.plotArea.axes
				.slice(0, axisIndex)
				.some((other) => other.kind === axis.kind);
			return { kind: axis.kind, vertical: vertical(axis, space), secondary };
		},
		seriesName: (seriesIndex) => seriesName(entries[seriesIndex]?.series, seriesIndex),
	};
}

/** The captioned picker entries of `space`. */
export const chartSpaceTargetOptions = (space: ChartSpace): ChartFormatTargetOption[] =>
	chartFormatTargetOptions(chartSpaceTargets(space), chartSpaceDescribe(space));
