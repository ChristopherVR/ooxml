import {
	setChartGrouping,
	setChartType,
	setChartSeriesType,
	type ChartTypeTarget,
} from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

const TYPES: Readonly<Record<string, ChartTypeTarget>> = {
	bar: { kind: 'bar', is3D: false },
	bar3D: { kind: 'bar', is3D: true },
	line: { kind: 'line', is3D: false },
	line3D: { kind: 'line', is3D: true },
	area: { kind: 'area', is3D: false },
	area3D: { kind: 'area', is3D: true },
	pie: { kind: 'pie', is3D: false },
	pie3D: { kind: 'pie', is3D: true },
	doughnut: { kind: 'doughnut', is3D: false },
	ofPie: { kind: 'ofPie', is3D: false },
	scatter: { kind: 'scatter', is3D: false },
	bubble: { kind: 'bubble', is3D: false },
	radar: { kind: 'radar', is3D: false },
	stock: { kind: 'stock', is3D: false },
	surface: { kind: 'surface', is3D: true },
};

export function planNeutralChartType(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const defaultType = probe.chartType;
	const wholeChanged = probe.chartType !== before.chartType;
	if (probe.chartType !== before.chartType) {
		const target = TYPES[probe.chartType];
		if (!target) return;
		const to = {
			...target,
			...(probe.chartType === 'surface' ? { is3D: !probe.surfaceTopView } : {}),
			...(probe.barDirection === undefined ? {} : { barDirection: probe.barDirection }),
			...(probe.grouping === undefined ? {} : { grouping: probe.grouping }),
		};
		edits.push((space) => setChartType(space, to));
		probe.chartType = before.chartType;
	}
	if (probe.series.length === before.series.length)
		probe.series.forEach((series, index) => {
			const old = before.series[index]!;
			if (
				series.seriesChartType === old.seriesChartType &&
				!(wholeChanged && series.seriesChartType !== undefined)
			)
				return;
			const target = TYPES[series.seriesChartType ?? defaultType];
			if (!target) return;
			const to = {
				...target,
				...(target.kind === 'bar' && probe.barDirection
					? { barDirection: probe.barDirection }
					: {}),
			};
			edits.push((space) => setChartSeriesType(space, index, to));
			if (old.seriesChartType === undefined) delete series.seriesChartType;
			else series.seriesChartType = old.seriesChartType;
		});
	if (probe.grouping !== before.grouping && probe.grouping !== undefined) {
		const grouping = probe.grouping;
		edits.push((space) => setChartGrouping(space, grouping));
		if (before.grouping === undefined) delete probe.grouping;
		else probe.grouping = before.grouping;
	}
}
