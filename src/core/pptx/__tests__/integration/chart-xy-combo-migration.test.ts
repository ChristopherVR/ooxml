import { describe, expect, it } from 'vitest';

import { parseChartSpace } from '../../../chart/index';
import { classicChartDecks } from './chart-save-parity';
import { saveFamilyEdit, type FamilyEdit } from './chart-family-migration';
import type { PptxChartData } from '../../index';

/**
 * One series of a vertical 2-D column, line or area chart turned into a scatter series, saved
 * through the neutral writer and through the object-tree writers over every committed deck. The
 * object-tree writers put the scatter series into a `c:scatterChart` with the `c:cat` and `c:val`
 * of the series it was, which CT_ScatterSer does not allow; the neutral writer writes `c:xVal` and
 * `c:yVal` on the same category and value axes, the structure Excel 16 saves for this edit. Both
 * reload with the same series types, names and values.
 */

const DECKS = await classicChartDecks();

const comboCandidate = (chart: PptxChartData) =>
	['bar', 'line', 'area'].includes(chart.chartType) &&
	chart.barDirection !== 'bar' &&
	chart.series.length > 1 &&
	!chart.series.some((series) => series.seriesChartType !== undefined);

const EDIT: FamilyEdit = {
	name: 'series to scatter',
	applies: comboCandidate,
	edit: (chart) => {
		chart.series[1]!.seriesChartType = 'scatter';
	},
	// By name: the neutral writer joins the column series around the scatter series into one group
	// (as Excel writes it) where the object-tree writers keep a column group on each side, so the
	// reloaded document order differs while every series keeps its own `c:order`.
	read: (chart) =>
		chart.series
			.map((series) => [series.name, series.seriesChartType ?? null, series.values] as const)
			.sort((a, b) => a[0].localeCompare(b[0])),
};

const scatterGroups = (xml: string) =>
	parseChartSpace(xml).chartSpace.plotArea.groups.filter((group) => group.kind === 'scatter');

let parts = 0;
let neutralParts = 0;

describe('series to scatter: object-tree writers vs neutral writer', () => {
	it.each(DECKS)('$name', async (deck) => {
		const [neutral, legacy] = await Promise.all([
			saveFamilyEdit(deck, EDIT, 'neutral'),
			saveFamilyEdit(deck, EDIT, 'legacy'),
		]);
		expect(neutral.length).toBe(legacy.length);
		neutral.forEach((saved, i) => {
			const old = legacy[i]!;
			parts++;
			expect(EDIT.read(saved.chart)).toEqual(EDIT.read(old.chart));
			// The object-tree writers keep the category-chart children in the scatter series.
			expect(scatterGroups(old.xml)[0]?.series[0]?.categories).toBeDefined();
			const [scatter] = scatterGroups(saved.xml);
			if (!scatter?.series[0]?.yValues) return;
			neutralParts++;
			expect(scatter.series[0].categories).toBeUndefined();
			const shared = parseChartSpace(saved.xml).chartSpace.plotArea.groups[0]!.axisIds;
			expect(scatter.axisIds).toEqual(shared);
		});
	});

	it('records the parts compared', () => {
		// Every applicable part, and how many the neutral plan saved (two splits inside a stack and two
		// groups with an unknown extension are refused and keep the object-tree path).
		expect({ parts, neutralParts }).toEqual({ parts: 16, neutralParts: 12 });
	});
});
