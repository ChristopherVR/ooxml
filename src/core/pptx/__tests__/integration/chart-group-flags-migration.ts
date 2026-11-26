import { describeFamilyMigrationEdits, type FamilyEdit } from './chart-family-migration';

/**
 * Group settings edited on their own (`chart-neutral-save-group-flags.ts`) saved through the
 * neutral writer and the object-tree writers over every committed deck. The object-tree writers
 * wrote no `c:varyColors`, `c:barDir` or `c:scatterStyle` on an existing part, so those edits were
 * lost on save (`legacyKept` 0); the neutral save keeps every one.
 */

const EDITS: FamilyEdit[] = [
	{
		name: 'vary colours',
		applies: (chart) =>
			chart.chartType !== 'combo' && !['stock', 'surface'].includes(chart.chartType),
		edit: (chart) => {
			chart.varyColors = !chart.varyColors;
		},
		read: (chart) => chart.varyColors ?? false,
	},
	{
		name: 'bar direction',
		applies: (chart) =>
			['bar', 'bar3D'].includes(chart.chartType) &&
			chart.series.every((series) => series.seriesChartType === undefined),
		edit: (chart) => {
			chart.barDirection = chart.barDirection === 'bar' ? 'col' : 'bar';
		},
		read: (chart) => chart.barDirection ?? 'col',
	},
	{
		name: 'scatter style',
		applies: (chart) => chart.chartType === 'scatter',
		edit: (chart) => {
			chart.scatterStyle = chart.scatterStyle === 'smoothMarker' ? 'lineMarker' : 'smoothMarker';
		},
		read: (chart) => chart.scatterStyle,
	},
	{
		name: 'standard grouping',
		applies: (chart) =>
			// A 3-D line or area cannot be clustered: both writers keep `standard` there.
			chart.grouping === 'clustered' && chart.chartType === 'bar3D',
		edit: (chart) => {
			chart.groupingStandard = !chart.groupingStandard;
		},
		read: (chart) => chart.groupingStandard ?? false,
	},
];

/**
 * Parts per edit, how many the object-tree save reloaded with the edit, and how many both paths
 * saved to the same canonical XML. Apart from the edited element (absent from every object-tree
 * save but the grouping one), the object-tree parts differ in what those writers do to the rest of
 * an edited part: they pin the theme title font (`+mn-lt` becomes `Aptos`), rewrite caches from
 * JavaScript numbers (`4.4000000000000004` becomes `4.4`), renumber line series `c:idx`, keep a
 * series `c:shape` before `c:cat`, and drop the switched-off `c:dLbls` of 16 pie charts.
 */
const EXPECTED = {
	'vary colours': { parts: 70, legacyKept: 0, xml: 0 },
	'bar direction': { parts: 31, legacyKept: 0, xml: 0 },
	'scatter style': { parts: 2, legacyKept: 0, xml: 0 },
	'standard grouping': { parts: 8, legacyKept: 8, xml: 1 },
};

export function describeGroupFlags(...names: (keyof typeof EXPECTED)[]): Promise<void> {
	return describeFamilyMigrationEdits(EDITS, EXPECTED, names);
}
