import type { PptxChartData } from '../../index';
import { classicChartDecks } from './chart-save-parity';
import { describeFamilyMigration, type FamilyEdit } from './chart-family-migration';

/**
 * Chart Filters edits (`chart-neutral-save-filtered.ts`) saved through the neutral writer and the
 * object-tree writers over every committed deck. The edits follow the shared UI's
 * `hideChartSeries` and `restoreFilteredSeries` (`ooxml-ui/pptx`, which the core does not import).
 */

const DECKS = await classicChartDecks();

const read = (chart: PptxChartData) => ({
	visible: chart.series.map((series) => [series.idx, series.name]),
	hidden: (chart.filteredSeries ?? []).map((entry) => [entry.idx, entry.name]),
});

const singleGroup = (chart: PptxChartData) =>
	chart.chartType !== 'combo' &&
	chart.series.every((series) => series.seriesChartType === undefined) &&
	new Set(chart.series.map((series) => series.axisId ?? 'primary')).size === 1;

const EDITS: FamilyEdit[] = [
	{
		name: 'hide a series',
		applies: (chart) =>
			chart.series.length >= 2 &&
			singleGroup(chart) &&
			chart.series.every((series) => series.idx !== undefined),
		edit: (chart) => {
			const series = chart.series[1]!;
			chart.filteredSeries = [
				...(chart.filteredSeries ?? []),
				{
					idx: series.idx!,
					order: series.idx!,
					name: series.name,
					categories: chart.categories,
					values: series.values,
					...(series.uniqueId ? { uniqueId: series.uniqueId } : {}),
					originalSeries: series,
				},
			];
			chart.series = chart.series.filter((_, i) => i !== 1);
		},
		read,
	},
	{
		name: 'restore a hidden series',
		applies: (chart) => (chart.filteredSeries?.length ?? 0) > 0 && singleGroup(chart),
		edit: (chart) => {
			const [entry, ...rest] = chart.filteredSeries!;
			const insertAt = chart.series.findIndex((series, i) => (series.idx ?? i) > entry!.idx);
			const restored = {
				name: entry!.name ?? `Series ${chart.series.length + 1}`,
				values: entry!.values ?? [],
				idx: entry!.idx,
				...(entry!.uniqueId ? { uniqueId: entry!.uniqueId } : {}),
			};
			chart.series.splice(insertAt < 0 ? chart.series.length : insertAt, 0, restored);
			chart.filteredSeries = rest.length > 0 ? rest : undefined;
		},
		read,
	},
	{
		name: 'drop a hidden series',
		applies: (chart) => (chart.filteredSeries?.length ?? 0) > 0,
		edit: (chart) => {
			chart.filteredSeries = chart.filteredSeries!.slice(1);
		},
		read,
	},
];

/**
 * The neutral save moves a hidden series into the extension with its references and formatting,
 * under the family wrapper Excel 16 writes for `Series.IsFiltered` (`c15:filteredBarSeries` for a
 * 3-D column too) and with `c15:formulaRef` in place of `c:f`; the object-tree writers rebuilt the
 * entry from literal caches (`c:numLit`, `c:strLit`) under a wrapper named after the group element
 * (`c15:filteredBar3DSeries`, which Excel does not write) and reloaded with the edit on 15 of the 40
 * hiding parts and neither restored part.
 */
describeFamilyMigration(DECKS, EDITS, {
	'hide a series': { parts: 40, legacyKept: 15, xml: 0 },
	'restore a hidden series': { parts: 2, legacyKept: 0, xml: 0 },
	'drop a hidden series': { parts: 2, legacyKept: 2, xml: 2 },
});
