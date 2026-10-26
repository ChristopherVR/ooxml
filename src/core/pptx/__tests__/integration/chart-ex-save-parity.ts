import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';

import * as neutral from '../../../chart/index';
import type { ChartExSpace } from '../../../chart/index';
import { CHART_EX_FIXTURES, chartExParts } from '../../../chart/ex/test-support/fixtures';
import * as pptx from '../../core/builders/sdk/chart-operations';
import { PptxHandler } from '../../core/PptxHandler';
import type { ChartPptxElement, PptxChartData, PptxElement } from '../../index';

/**
 * Support for the ChartEx save harness (`chart-ex-save-parity.test.ts`), the chartex counterpart of
 * `chart-save-parity.ts`: decks holding every committed chartex part (the gallery deck as it is,
 * and with the Word and Excel parts swapped into its chartex parts), one PowerPoint edit on every
 * chart it applies to, saved through `PptxHandler.save` and through the neutral reference
 * (`parseChartExSpace`, the `chart/edit` chartex operation, `writeChartExSpace`).
 */

const GALLERY = path.resolve(__dirname, '../fixtures/e2e/chart-gallery.pptx');

export interface ChartExPart {
	part: string;
	space: ChartExSpace;
}

export interface ChartExDeck {
	name: string;
	bytes: Uint8Array;
	parts: ChartExPart[];
}

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
			const parts: ChartExPart[] = [];
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
			return {
				name: index === 0 ? 'chart-gallery.pptx' : `docx/xlsx parts ${String(index)}`,
				bytes: await zip.generateAsync({ type: 'uint8array' }),
				parts,
			};
		}),
	);
}

export interface ChartExSaveEdit {
	name: string;
	applies?: (space: ChartExSpace, chart: PptxChartData) => boolean;
	pptx: (element: ChartPptxElement) => void;
	neutral: (space: ChartExSpace, chart: PptxChartData) => ChartExSpace;
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

async function load(bytes: Uint8Array) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice().buffer as ArrayBuffer);
	const charts = new Map<string, { element: ChartPptxElement; slide: { isDirty?: boolean } }>();
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

export interface SavedChartEx {
	xml: string;
	chart: PptxChartData;
}

async function readBack(bytes: Uint8Array, parts: readonly string[]) {
	const zip = await JSZip.loadAsync(bytes);
	const loaded = await load(bytes);
	const out = new Map<string, SavedChartEx>();
	for (const part of parts)
		out.set(part, {
			xml: await zip.file(part)!.async('string'),
			chart: loaded.charts.get(part)!.element.chartData!,
		});
	return out;
}

export interface ChartExTarget extends ChartExPart {
	chart: PptxChartData;
}

/** The parts of `deck` the edit applies to, with the chart as PowerPoint loaded it. */
export async function chartExTargets(
	deck: ChartExDeck,
	edit: ChartExSaveEdit,
): Promise<ChartExTarget[]> {
	const loaded = await load(deck.bytes);
	return deck.parts.flatMap((part) => {
		const chart = loaded.charts.get(part.part)?.element.chartData;
		return chart !== undefined && (edit.applies?.(part.space, chart) ?? true)
			? [{ ...part, chart: structuredClone(chart) }]
			: [];
	});
}

/** PowerPoint's save of the edited charts. */
export async function chartExSavedByPptx(
	deck: ChartExDeck,
	edit: ChartExSaveEdit,
	targets: readonly ChartExTarget[],
) {
	const loaded = await load(deck.bytes);
	for (const { part } of targets) {
		const target = loaded.charts.get(part)!;
		edit.pptx(target.element);
		target.slide.isDirty = true;
	}
	const saved = await loaded.handler.save(loaded.data.slides);
	return readBack(
		saved,
		targets.map((t) => t.part),
	);
}

/** The neutral reference: edit each parsed part and write it. */
export async function chartExSavedByNeutral(
	deck: ChartExDeck,
	edit: ChartExSaveEdit,
	targets: readonly ChartExTarget[],
) {
	const zip = await JSZip.loadAsync(deck.bytes);
	for (const { part, space, chart } of targets)
		zip.file(part, neutral.writeChartExSpace(edit.neutral(space, chart)));
	return readBack(
		await zip.generateAsync({ type: 'uint8array' }),
		targets.map((t) => t.part),
	);
}
