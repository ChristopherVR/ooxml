// Pareto and region map creation: the markup Office writes for new charts, the shared painter's
// rendering of them, and the edits that keep them consistent.
import { describe, expect, it } from 'vitest';
import { createChartExSpace } from './create';
import { parseChartExSpace } from './parse';
import { writeChartExSpace } from './write';
import { chartExDataLayout } from './data-layout';
import { applyChartExSpaceDataSheetEdits, chartExSpaceDataSheet } from '../edit/chartex-sheet';
import { applyChartExEditAction } from '../edit/chartex-action';
import { addChartExSeries } from '../edit/chartex-structure';
import { chartExFamilyOf, isChartExPareto } from '../edit/chartex-family';
import { chartExSummaryFromModel } from '../render/summary-from-chartex';
import { chartSummaryView } from '../render/chart-view';
import { chartThemePalette, renderChartSummarySvg } from '../render/render-space';
import { paretoCumulative, paretoOrder } from '../pareto';

const causes = {
	categories: ['Late', 'Damaged', 'Wrong item', 'Missing'],
	series: [{ name: 'Count', values: [42, 18, 27, 9] }],
};

describe('pareto charts', () => {
	it("writes Excel's pareto markup: aggregation, an owned paretoLine and a percentage axis", () => {
		const xml = writeChartExSpace(createChartExSpace(causes, 'pareto'));
		expect(xml).toContain('<cx:layoutPr><cx:aggregation/></cx:layoutPr><cx:axisId val="1"/>');
		expect(xml).toMatch(
			/<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId val="2"\/><\/cx:series>/,
		);
		expect(xml).toContain('<cx:catScaling gapWidth="0"/>');
		expect(xml).toContain('<cx:valScaling max="1" min="0"/><cx:units unit="percentage"/>');
		const parsed = parseChartExSpace(xml);
		expect(parsed.issues).toEqual([]);
		expect(chartExFamilyOf(parsed.chartExSpace)).toBe('pareto');
	});

	it('orders bars descending and draws the cumulative share on a percent axis', () => {
		expect(paretoOrder([42, 18, 27, 9])).toEqual([0, 2, 1, 3]);
		expect(paretoCumulative([2, 1, 1])).toEqual([0.5, 0.75, 1]);
		const summary = chartExSummaryFromModel(createChartExSpace(causes, 'pareto')).summary!;
		expect(summary.histogram?.pareto).toEqual({});
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		expect(view.categories).toEqual(['Late', 'Wrong item', 'Damaged', 'Missing']);
		expect(view.series[0]!.values).toEqual([42, 27, 18, 9]);
		const line = view.extraGroups?.[0];
		expect(line).toMatchObject({ type: 'line', secondary: true });
		expect(line!.series[0]!.values.at(-1)).toBe(1);
		expect(view.secondaryAxis).toMatchObject({ min: 0, max: 1, percent: true });
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		expect(svg).toContain('100%');
	});

	it('keeps the data sheet to the data series and refuses a second series', () => {
		const space = createChartExSpace(causes, 'pareto');
		expect(chartExSpaceDataSheet(space).series).toHaveLength(1);
		const edited = applyChartExSpaceDataSheetEdits(space, [
			{ kind: 'value', series: 0, category: 3, value: 50 },
			{ kind: 'series-name', series: 0, name: 'Returns' },
		]);
		expect(chartExSpaceDataSheet(edited).series[0]).toEqual({
			name: 'Returns',
			values: [42, 18, 27, 50],
		});
		expect(isChartExPareto(edited)).toBe(true);
		expect(() => addChartExSeries(space, 'Second')).toThrow(/one data series/);
		const { rows, space: laidOut } = chartExDataLayout(space, 'Sheet1');
		expect(rows[0]).toEqual(['Data 1', 'Data 2', 'Count']);
		expect(laidOut.plotArea.series[1]!.tx).toBeUndefined();
	});

	it('switches into and out of pareto through Change Chart Type', () => {
		const histogram = applyChartExEditAction(createChartExSpace(causes, 'pareto'), {
			op: 'layout',
			layout: 'clusteredColumn',
		});
		expect(histogram.plotArea.series.map((s) => s.layoutId)).toEqual(['clusteredColumn']);
		expect(histogram.plotArea.axes.map((a) => a.id)).toEqual([0, 1]);
		const back = applyChartExEditAction(
			applyChartExEditAction(histogram, { op: 'layout', layout: 'funnel' }),
			{ op: 'layout', layout: 'pareto' },
		);
		expect(chartExFamilyOf(back)).toBe('pareto');
		expect(back.plotArea.series[0]!.layoutProperties).toEqual({ aggregation: true });
	});
});

describe('region maps', () => {
	const places = {
		categories: ['France', 'Brazil', 'Canada'],
		series: [{ name: 'Visitors', values: [30, 60, 90] }],
	};

	it('writes the schema geography without a provider cache', () => {
		const space = createChartExSpace(places, 'regionMap', {
			geography: { cultureLanguage: 'fr-FR', cultureRegion: 'FR', projectionType: 'robinson' },
		});
		const xml = writeChartExSpace(space);
		expect(xml).toContain('<cx:numDim type="colorVal">');
		expect(xml).toContain(
			'<cx:geography projectionType="robinson" cultureLanguage="fr-FR" cultureRegion="FR" attribution=""/>',
		);
		expect(xml).not.toContain('geoCache');
		expect(parseChartExSpace(xml).issues).toEqual([]);
	});

	it('paints bundled approximate outlines and says so', () => {
		const summary = chartExSummaryFromModel(createChartExSpace(places, 'regionMap')).summary!;
		const svg = renderChartSummarySvg(summary, {
			width: 480,
			height: 300,
			mapApproximationLabel: 'Approximate outlines',
		});
		expect(svg).toContain('data-map-geometry="approximate"');
		expect(svg).toContain('Approximate outlines');
		expect(svg).toMatch(/<title>France: 30<\/title>/);
	});

	it('edits projection, area and labels with the schema values only', () => {
		const space = createChartExSpace(places, 'regionMap');
		const next = applyChartExEditAction(space, {
			op: 'regionMap',
			seriesIndex: 0,
			options: {
				projectionType: 'miller',
				viewedRegionType: 'world',
				regionLabelLayout: 'showAll',
			},
		});
		const layout = next.plotArea.series[0]!.layoutProperties!;
		expect(layout.regionLabelLayout).toBe('showAll');
		expect(layout.geography).toMatchObject({ projectionType: 'miller', viewedRegionType: 'world' });
		expect(() =>
			applyChartExEditAction(space, {
				op: 'regionMap',
				seriesIndex: 0,
				options: { projectionType: 'globe' as 'miller' },
			}),
		).toThrow(/projection/);
		const auto = applyChartExEditAction(next, {
			op: 'regionMap',
			seriesIndex: 0,
			options: { projectionType: 'auto' },
		});
		expect(auto.plotArea.series[0]!.layoutProperties!.geography!.projectionType).toBeUndefined();
	});
});
