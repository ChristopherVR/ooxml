import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import * as neutral from '../../../chart/index';
import { parseChartSpace, writeChartSpace, type ChartSpace } from '../../../chart/index';
import * as pptx from '../../core/builders/sdk/chart-operations';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxElement } from '../../core/types/elements';

/**
 * Adapter proof for the shared chart edit core (docs/agnostic-core-plan.md, chart step C1): the
 * same edit made by a neutral operation (`chart/edit`, on `parseChartSpace` of the part, written
 * back with `writeChartSpace`) and by PowerPoint's operation (`builders/sdk/chart-operations`, on
 * the loaded `PptxChartData`, saved by the pptx writers) must load back, through the existing
 * `chart-from-neutral` adapter, as the same `PptxChartData`. The whole model is compared; where
 * the two editors legitimately differ, the difference is listed per operation, so a change on
 * either side turns the suite red.
 */

const DECK = path.resolve(__dirname, '../fixtures/e2e/chart-gallery.pptx');
const PART = /^ppt\/charts\/chart\d+\.xml$/u;

// A copy through the constructor: `slice()` on a Node Buffer (`readFileSync`) is a view, not a copy.
const toBuffer = (bytes: Uint8Array): ArrayBuffer => new Uint8Array(bytes).buffer;

type Slide = { isDirty?: boolean };

async function load(bytes: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes);
	const charts = new Map<string, { element: ChartPptxElement; slide: Slide }>();
	for (const slide of data.slides) {
		const visit = (list: PptxElement[] | undefined) => {
			for (const element of list ?? []) {
				const chart = element as ChartPptxElement;
				if (element.type === 'chart' && chart.chartData?.chartPartPath)
					charts.set(chart.chartData.chartPartPath, { element: chart, slide });
				visit((element as { children?: PptxElement[] }).children);
			}
		};
		visit(slide.elements);
	}
	return { charts, handler, data };
}

const source = readFileSync(DECK);
const zip = await JSZip.loadAsync(source);
const original = await load(toBuffer(source));
/** Classic bar and line charts with one group: the families both editors handle the same way. */
const CANDIDATES: string[] = [];
for (const part of Object.keys(zip.files).sort()) {
	if (!PART.test(part) || !original.charts.has(part)) continue;
	const { chartSpace, issues } = parseChartSpace(await zip.file(part)!.async('string'));
	const groups = chartSpace.plotArea.groups;
	if (issues.length === 0 && groups.length === 1 && ['bar', 'line'].includes(groups[0]!.kind))
		CANDIDATES.push(part);
}

/** Neutral path: edit the parsed part, write it into the deck, load the deck. */
async function neutralEdit(part: string, edit: (space: ChartSpace) => ChartSpace) {
	const copy = await JSZip.loadAsync(source);
	const { chartSpace } = parseChartSpace(await copy.file(part)!.async('string'));
	copy.file(part, writeChartSpace(edit(chartSpace)));
	const loaded = await load(toBuffer(await copy.generateAsync({ type: 'uint8array' })));
	return loaded.charts.get(part)!.element.chartData!;
}

/** PowerPoint path: edit the loaded model, save the deck, load it again. */
async function pptxEdit(part: string, edit: (element: ChartPptxElement) => void) {
	const loaded = await load(toBuffer(source));
	const target = loaded.charts.get(part)!;
	edit(target.element);
	target.slide.isDirty = true;
	const saved = await loaded.handler.save(loaded.data.slides);
	return (await load(toBuffer(saved))).charts.get(part)!.element.chartData!;
}

/** Paths where two JSON values differ, with both values. */
function differences(a: unknown, b: unknown, at = ''): string[] {
	if (JSON.stringify(a) === JSON.stringify(b)) return [];
	if (a && b && typeof a === 'object' && typeof b === 'object') {
		const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
		return [...keys].flatMap((key) =>
			differences(
				(a as Record<string, unknown>)[key],
				(b as Record<string, unknown>)[key],
				`${at}.${key}`,
			),
		);
	}
	return [`${at}: neutral ${JSON.stringify(a)}, pptx ${JSON.stringify(b)}`];
}

const red = { kind: 'srgb' as const, value: 'FF0000', transforms: [] };
type Chart = NonNullable<ChartPptxElement['chartData']>;

interface Case {
	name: string;
	neutral: (space: ChartSpace) => ChartSpace;
	pptx: (element: ChartPptxElement) => void;
	/** What the neutral edit must have produced, so an edit that does nothing cannot pass. */
	check: (chart: Chart) => void;
	/** Differences the two editors are known to have, given the chart before the edit. */
	expected?: (before: Chart) => string[];
}

const CASES: Case[] = [
	{
		name: 'title',
		neutral: (space) => neutral.setChartTitle(space, 'Edited title'),
		pptx: (element) => pptx.setChartTitle(element, 'Edited title'),
		check: (chart) => expect(chart.title).toBe('Edited title'),
	},
	{
		name: 'legend position',
		neutral: (space) => neutral.setChartLegend(space, { position: 't' }),
		pptx: (element) => pptx.setChartLegend(element, { position: 't' }),
		check: (chart) => expect(chart.style).toMatchObject({ hasLegend: true, legendPosition: 't' }),
	},
	{
		name: 'legend hidden',
		neutral: (space) => neutral.setChartLegend(space, { visible: false }),
		pptx: (element) => pptx.setChartLegend(element, { show: false }),
		check: (chart) => expect(chart.style?.hasLegend).toBeUndefined(),
	},
	{
		name: 'series values and categories',
		neutral: (space) =>
			neutral.setChartCategories(neutral.setChartSeriesValues(space, 0, [7, 14, 21]), [
				'A',
				'B',
				'C',
			]),
		pptx: (element) => {
			pptx.updateChartSeriesValues(element, 0, [7, 14, 21]);
			pptx.setChartCategories(element, ['A', 'B', 'C']);
		},
		check: (chart) => {
			expect(chart.series[0]!.values).toEqual([7, 14, 21]);
			expect(chart.categories).toEqual(['A', 'B', 'C']);
		},
	},
	{
		name: 'series name',
		neutral: (space) => neutral.renameChartSeries(space, 0, 'Renamed'),
		pptx: (element) => {
			element.chartData!.series[0]!.name = 'Renamed';
		},
		check: (chart) => expect(chart.series[0]!.name).toBe('Renamed'),
	},
	{
		name: 'data labels',
		neutral: (space) => neutral.setChartDataLabels(space, { show: true, position: 'ctr' }),
		pptx: (element) => pptx.setChartDataLabels(element, { show: true, position: 'ctr' }),
		check: (chart) => expect(chart.style?.hasDataLabels).toBe(true),
	},
	{
		name: 'value axis scaling',
		neutral: (space) => neutral.setChartAxis(space, 'val', { min: 0, max: 99, majorUnit: 11 }),
		pptx: (element) => pptx.setChartAxis(element, 'valAx', { min: 0, max: 99, majorUnit: 11 }),
		check: (chart) =>
			expect(chart.axes?.find((axis) => axis.axisType === 'valAx')).toMatchObject({
				min: 0,
				max: 99,
				majorUnit: 11,
			}),
	},
	{
		name: 'series colour',
		neutral: (space) => neutral.setChartSeriesColor(space, 0, red),
		pptx: (element) => pptx.setChartSeriesColor(element, 0, '#FF0000'),
		check: (chart) => expect(chart.series[0]!.color?.toUpperCase()).toBe('#FF0000'),
	},
	{
		name: 'type switch to line',
		neutral: (space) => neutral.setChartType(space, { kind: 'line' }),
		pptx: (element) => pptx.setChartType(element, 'line'),
		check: (chart) => expect(chart.chartType).toBe('line'),
		// Both paths now use the shared type operation, including its authored line width/cap.
	},
];

describe('neutral chart edits load as PowerPoint chart edits', () => {
	it('covers the bar and line charts of the gallery deck', () => {
		expect(CANDIDATES).toEqual([
			'ppt/charts/chart1.xml',
			'ppt/charts/chart10.xml',
			'ppt/charts/chart2.xml',
			'ppt/charts/chart9.xml',
		]);
	});

	describe.each(CANDIDATES)('%s', (part) => {
		it.each(CASES)('$name', async (item) => {
			const fromNeutral = await neutralEdit(part, item.neutral);
			const fromPptx = await pptxEdit(part, item.pptx);
			item.check(fromNeutral);
			const before = original.charts.get(part)!.element.chartData!;
			expect(differences(fromNeutral, fromPptx).sort()).toEqual(
				(item.expected?.(before) ?? []).sort(),
			);
		});
	});
});
