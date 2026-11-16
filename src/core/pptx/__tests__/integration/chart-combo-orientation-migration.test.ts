import type { PptxChartData } from '../../index';
import { classicChartDecks } from './chart-save-parity';
import { describeFamilyMigration, type FamilyEdit } from './chart-family-migration';

/**
 * Per-series combinations on horizontal bars and on two value axes (`chart/edit/series-type.ts`)
 * saved through the neutral writer and the object-tree writers over every committed deck. The
 * neutral save writes the structures Excel 16 saves for `Series.ChartType`: a line series of a
 * horizontal bar chart on a secondary axis pair, and a line series of an XY chart on the x axis
 * turned into a category axis the scatter series share.
 */

const DECKS = await classicChartDecks();

const read = (chart: PptxChartData) =>
	chart.series.map((series) => series.seriesChartType ?? chart.chartType);

const plain = (chart: PptxChartData) =>
	chart.series.length >= 2 &&
	chart.series.every((series) => series.seriesChartType === undefined) &&
	new Set(chart.series.map((series) => series.axisId ?? 'primary')).size === 1;

const EDITS: FamilyEdit[] = [
	{
		name: 'line series on horizontal bars',
		applies: (chart) => chart.chartType === 'bar' && chart.barDirection === 'bar' && plain(chart),
		edit: (chart) => {
			chart.series[1]!.seriesChartType = 'line';
		},
		read: (chart) => read(chart).map((type) => (type === 'combo' ? 'bar' : type)),
	},
	{
		name: 'line series on two value axes',
		applies: (chart) => chart.chartType === 'scatter' && plain(chart),
		edit: (chart) => {
			chart.series[1]!.seriesChartType = 'line';
		},
		read: (chart) => read(chart).map((type) => (type === 'combo' ? 'scatter' : type)),
	},
];

/**
 * No committed deck holds a 2-D horizontal bar chart with two plain series, so that case is unit
 * evidence (`chart/edit/series-type.test.ts`) and native acceptance. On the 2 XY parts the
 * object-tree writers wrote the line series with `c:xVal` and `c:yVal`, which CT_LineSer does not
 * allow, on the two value axes; the neutral writer writes `c:cat` and `c:val` on a category axis.
 */
describeFamilyMigration(DECKS, EDITS, {
	'line series on horizontal bars': { parts: 0, legacyKept: 0, xml: 0 },
	'line series on two value axes': { parts: 2, legacyKept: 2, xml: 0 },
});
