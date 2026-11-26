import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';

import * as neutral from '../../../chart/index';
import type { ChartExSpace } from '../../../chart/index';
import { CHART_EX_FIXTURES, chartExParts } from '../../../chart/ex/test-support/fixtures';
import * as pptx from '../../core/builders/sdk/chart-operations';
import type { PptxChartData } from '../../index';
import type { ChartFixture, ChartFixtureEdit } from './chart-save-fixture';

/**
 * Support for the ChartEx save harness (`chart-ex-save-parity.test.ts`), the chartex counterpart of
 * `chart-save-parity.ts`: decks holding every committed chartex part (the gallery deck as it is,
 * and with the Word and Excel parts swapped into its chartex parts), one PowerPoint edit on every
 * chart it applies to, saved through `PptxHandler.save` and through the neutral reference
 * (`parseChartExSpace`, the `chart/edit` chartex operation, `writeChartExSpace`).
 */

const GALLERY = path.resolve(__dirname, '../fixtures/e2e/chart-gallery.pptx');

type ChartExDeck = ChartFixture<ChartExSpace>;

/** The gallery deck, and copies with the other committed chartex parts in its chartex slots. */
export async function chartExDecks(): Promise<ChartExDeck[]> {
	const gallery = readFileSync(GALLERY);
	const slots = (await chartExParts(CHART_EX_FIXTURES[0]![0])).map(([name]) => name);
	const others: string[] = [];
	for (const [file] of CHART_EX_FIXTURES.slice(1))
		for (const [, xml] of await chartExParts(file)) others.push(xml);
	const batches: string[][] = [[]];
	for (let i = 0; i < others.length; i += slots.length)
		batches.push(others.slice(i, i + slots.length));
	return Promise.all(
		batches.map(async (batch, index) => {
			const zip = await JSZip.loadAsync(gallery);
			batch.forEach((xml, slot) => zip.file(slots[slot]!, xml));
			const parts: { part: string; space: ChartExSpace }[] = [];
			for (const part of slots.slice(0, batch.length || slots.length)) {
				// This codec harness swaps parts across unrelated packages. Their original workbook
				// links are unavailable here; literal caches isolate parser/writer parity. Package
				// synchronization and refusal are covered by chart-ex-neutral-save.test.ts.
				const xml = (await zip.file(part)!.async('string'))
					.replace(/<cx:f\b[^>]*>[\s\S]*?<\/cx:f>/gu, '')
					.replace(/<cx:externalData\b[^>]*(?:\/>|>[\s\S]*?<\/cx:externalData>)/gu, '');
				zip.file(part, xml);
				parts.push({ part, space: neutral.parseChartExSpace(xml).chartExSpace });
			}
			const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
			return {
				name: index === 0 ? 'chart-gallery.pptx' : `docx/xlsx parts ${String(index)}`,
				bytes,
				parts,
				archive: await JSZip.loadAsync(bytes),
			};
		}),
	);
}

interface ChartExSaveEdit extends ChartFixtureEdit<ChartExSpace> {
	name: string;
}

const isHierarchy = (space: ChartExSpace) =>
	space.plotArea.series.some((s) => s.layoutId === 'treemap' || s.layoutId === 'sunburst');
const editedValues = (chart: PptxChartData) =>
	(chart.series[0]?.values ?? []).map((value) => value * 2 + 1);

export const CHART_EX_SAVE_EDITS: ChartExSaveEdit[] = [
	{ name: 'no edit (slide dirty)', pptx: () => {}, neutral: (space) => space },
	{
		name: 'title text',
		pptx: (element) => pptx.setChartTitle(element, 'Edited title'),
		neutral: (space) => neutral.setChartExTitle(space, 'Edited title'),
	},
	{
		name: 'title hidden',
		pptx: (element) => {
			element.chartData!.style = { ...element.chartData!.style, hasTitle: false };
		},
		neutral: (space) => neutral.setChartExTitle(space, undefined),
	},
	{
		name: 'legend position',
		pptx: (element) => pptx.setChartLegend(element, { position: 't' }),
		neutral: (space, chart) =>
			chart.style?.hasLegend === false
				? space
				: neutral.setChartExLegend(space, { visible: true, position: 't' }),
	},
	{
		name: 'legend hidden',
		pptx: (element) => pptx.setChartLegend(element, { show: false }),
		neutral: (space) => neutral.setChartExLegend(space, { visible: false }),
	},
	{
		name: 'series name',
		pptx: (element) => {
			element.chartData!.series[0]!.name = 'Renamed';
		},
		neutral: (space) => neutral.renameChartExSeries(space, 0, 'Renamed'),
	},
	{
		name: 'series values',
		pptx: (element) => pptx.updateChartSeriesValues(element, 0, editedValues(element.chartData!)),
		neutral: (space, chart) => neutral.setChartExSeriesValues(space, 0, editedValues(chart)),
	},
	{
		name: 'categories',
		applies: (space, chart) => !isHierarchy(space) && chart.categories.length > 0,
		pptx: (element) =>
			pptx.setChartCategories(
				element,
				element.chartData!.categories.map((_, i) => `Cat ${String(i + 1)}`),
			),
		neutral: (space, chart) =>
			neutral.setChartExCategories(
				space,
				chart.categories.map((_, i) => `Cat ${String(i + 1)}`),
			),
	},
	{
		name: 'data labels',
		pptx: (element) => pptx.setChartDataLabels(element, { show: true }),
		neutral: (space) =>
			neutral.applyChartExEditIntent(neutral.setChartExDataLabelsVisible(space, true), {
				op: 'dataLabels',
				edit: { showValue: true },
			}),
	},
];
