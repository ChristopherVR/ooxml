import { describe, expect, it } from 'vitest';

import { planNeutralChartSave } from '../../core/utils/chart-neutral-save-plan';
import type { PptxChartData } from '../../index';
import { classicChartDecks } from './chart-save-parity';
import { familyXmlDifferences, saveFamilyEdit, type FamilyEdit } from './chart-family-migration';

/**
 * Group subtype settings (`chart-neutral-save-subtype.ts`) saved through the neutral writer and
 * through the object-tree writers over every committed deck: both saves must reload with the edited
 * value, and the canonical part XML may differ only in what the object-tree writers do to the rest
 * of an edited part (recorded below).
 */

const DECKS = await classicChartDecks();

const EDITS: FamilyEdit[] = [
	{
		name: '3-D bar shape',
		applies: (chart) => chart.chartType === 'bar3D',
		edit: (chart) => {
			chart.barShape = chart.barShape === 'cylinder' ? 'pyramid' : 'cylinder';
		},
		read: (chart) => chart.barShape,
	},
	{
		name: '3-D gap depth',
		applies: (chart) => ['bar3D', 'area3D', 'line3D'].includes(chart.chartType),
		edit: (chart) => {
			chart.gapDepth = chart.gapDepth === 80 ? 220 : 80;
		},
		read: (chart) => chart.gapDepth,
	},
	{
		name: 'radar style',
		applies: (chart) => chart.chartType === 'radar',
		edit: (chart) => {
			chart.radarStyle = chart.radarStyle === 'filled' ? 'marker' : 'filled';
		},
		read: (chart) => chart.radarStyle,
	},
	{
		name: 'surface wireframe',
		applies: (chart) => chart.chartType === 'surface',
		edit: (chart) => {
			chart.wireframe = !chart.wireframe;
		},
		read: (chart) => chart.wireframe,
	},
	{
		name: 'pie of pie split',
		applies: (chart) => chart.chartType === 'ofPie' && chart.ofPieOptions !== undefined,
		edit: (chart) => {
			chart.ofPieOptions = {
				...chart.ofPieOptions!,
				splitType: 'pos',
				splitPos: 2,
				secondPieSize: 60,
				gapWidth: 120,
			};
		},
		read: (chart) => chart.ofPieOptions,
	},
	{
		name: 'bubble options',
		applies: (chart) => chart.chartType === 'bubble',
		edit: (chart) => {
			chart.bubbleOptions = {
				...chart.bubbleOptions,
				bubbleScale: 55,
				showNegativeBubbles: true,
				sizeRepresents: 'w',
			};
		},
		read: (chart) => chart.bubbleOptions,
	},
];

/**
 * Parts per edit, and how many saved to the same canonical XML both ways. Every other part differs
 * only in what the object-tree writers do to the rest of an edited part, none of it the edit: they
 * pin the theme title font (`+mn-lt` becomes `Aptos`), rewrite caches from JavaScript numbers
 * (`4.4000000000000004` becomes `4.4`), normalise manual layout numbers, write a resolved series
 * colour onto a radar or 3-D line series' outline (over an authored `a:noFill`), and keep the series
 * `c:shape` of `bar3d-picture-fill.pptx` before `c:cat`, where the neutral writer puts it in its
 * schema place after `c:val`. No fixture holds a
 * pie of pie chart, so that edit is covered by `chart-neutral-save-subtype.test.ts`.
 */
const EXPECTED: Record<string, { parts: number; xml: number }> = {
	'3-D bar shape': { parts: 11, xml: 1 },
	'3-D gap depth': { parts: 15, xml: 1 },
	'radar style': { parts: 2, xml: 0 },
	'surface wireframe': { parts: 3, xml: 1 },
	'pie of pie split': { parts: 0, xml: 0 },
	'bubble options': { parts: 2, xml: 2 },
};
const TALLY: Record<string, { parts: number; xml: number }> = {};

describe('chart group subtype: object-tree writers vs neutral writer', () => {
	describe.each(EDITS)('$name', (edit) => {
		it.each(DECKS)('$name', async (deck) => {
			const [neutral, legacy] = await Promise.all([
				saveFamilyEdit(deck, edit, 'neutral'),
				saveFamilyEdit(deck, edit, 'legacy'),
			]);
			expect(neutral.length).toBe(legacy.length);
			const tally = (TALLY[edit.name] ??= { parts: 0, xml: 0 });
			neutral.forEach((saved, i) => {
				const old = legacy[i]!;
				expect(saved.part).toBe(old.part);
				expect(edit.read(saved.chart)).toEqual(edit.read(old.chart));
				tally.parts++;
				const diff = familyXmlDifferences(old.xml, saved.xml);
				if (process.env['CHART_FAMILY_DETAIL'] === edit.name && diff.length > 0)
					console.log([`${deck.name} ${saved.part}`, ...diff].join('\n'));
				if (diff.length === 0) tally.xml++;
			});
		});
	});

	it('claims every edit on the neutral path', () => {
		const chart = (type: string, extra: Partial<PptxChartData>): PptxChartData =>
			({ chartType: type, categories: [], series: [], ...extra }) as PptxChartData;
		const cases: [PptxChartData, (c: PptxChartData) => void][] = [
			[chart('bar3D', { barShape: 'box' }), (c) => (c.barShape = 'cone')],
			[chart('line3D', {}), (c) => (c.gapDepth = 90)],
			[chart('radar', { radarStyle: 'marker' }), (c) => (c.radarStyle = 'filled')],
			[chart('surface', { wireframe: false }), (c) => (c.wireframe = true)],
			[
				chart('ofPie', { ofPieOptions: { ofPieType: 'pie' } }),
				(c) => (c.ofPieOptions = { ofPieType: 'bar', splitType: 'percent', splitPos: 10 }),
			],
			[chart('bubble', {}), (c) => (c.bubbleOptions = { bubbleScale: 40 })],
		];
		for (const [before, edit] of cases) {
			const after = structuredClone(before);
			edit(after);
			expect(planNeutralChartSave(before, after)).toBeDefined();
		}
	});

	it('records the parts compared', () => {
		if (process.env['CHART_FAMILY_REPORT']) console.log(JSON.stringify(TALLY, null, 1));
		expect(TALLY).toEqual(EXPECTED);
	});
});
