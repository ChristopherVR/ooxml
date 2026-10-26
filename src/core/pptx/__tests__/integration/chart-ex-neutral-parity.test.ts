import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { parseChartExSpace } from '../../../chart/index';
import { CHART_EX_FIXTURES, chartExParts } from '../../../chart/ex/test-support/fixtures';
import { FULL } from '../../../chart/ex/test-support/full-part';
import { PptxHandler } from '../../index';
import type { PptxChartData, PptxElement } from '../../index';
import { projectNeutralChartEx, projectPptxChartEx } from './chart-ex-neutral-parity-project';

/**
 * ChartEx parity harness (docs/agnostic-core-plan.md, chart step C5): every committed chartex part
 * (the PowerPoint chart gallery, the Word and Excel chart-family samples) and a synthetic part using
 * every modelled element is read through BOTH public models: `PptxHandler` (the host styles and
 * shared ChartEx read adapter, `PptxChartData`) and `parseChartExSpace` (the shared
 * DOM, `ChartExSpace`) - and both are projected onto one comparison shape
 * (`chart-ex-neutral-parity-project.ts`) that must be equal. Parts from Word and Excel are swapped
 * into the gallery deck's chartex parts so the handler reads them in place.
 */

const GALLERY = path.resolve(__dirname, '../fixtures/e2e/chart-gallery.pptx');

/** Edge parts: bare empty elements fast-xml-parser reads as empty strings. */
const NS =
	'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const EDGE = [
	`<cx:chartSpace ${NS}><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:dataId val="0"/><cx:layoutPr><cx:aggregation/></cx:layoutPr></cx:series><cx:series layoutId="paretoLine" ownerIdx="0"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:units unit="thousands"/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend/></cx:chart></cx:chartSpace>`,
];

async function corpus(): Promise<string[]> {
	const parts: string[] = [];
	for (const [file] of CHART_EX_FIXTURES)
		for (const [, xml] of await chartExParts(file)) parts.push(xml);
	return [...parts, FULL, ...EDGE];
}

/** The gallery deck with its chartex parts replaced by `parts` (in slot order). */
async function deckWith(parts: string[], slots: string[]): Promise<Buffer> {
	const zip = await JSZip.loadAsync(readFileSync(GALLERY));
	parts.forEach((xml, index) => zip.file(slots[index]!, xml));
	return zip.generateAsync({ type: 'nodebuffer' });
}

async function chartsByPart(bytes: Buffer): Promise<Map<string, PptxChartData>> {
	const data = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
	);
	const byPart = new Map<string, PptxChartData>();
	const visit = (list: PptxElement[] | undefined) => {
		for (const element of list ?? []) {
			const chartData = (element as { chartData?: PptxChartData }).chartData;
			if (element.type === 'chart' && chartData?.chartPartPath)
				byPart.set(chartData.chartPartPath, chartData);
			visit((element as { children?: PptxElement[] }).children);
		}
	};
	for (const slide of data.slides) visit(slide.elements);
	return byPart;
}

const PARTS = await corpus();
const SLOTS = (await chartExParts('../../pptx/__tests__/fixtures/e2e/chart-gallery.pptx')).map(
	([name]) => name,
);
const BATCHES: string[][] = [];
for (let i = 0; i < PARTS.length; i += SLOTS.length) BATCHES.push(PARTS.slice(i, i + SLOTS.length));

describe('chartex parity: PowerPoint consumer vs neutral parseChartExSpace', () => {
	it.each(BATCHES.map((batch, index) => [index, batch] as const))(
		'batch %i projects every part identically',
		async (_index, batch) => {
			const byPart = await chartsByPart(await deckWith(batch, SLOTS));
			batch.forEach((xml, slot) => {
				const pptx = byPart.get(SLOTS[slot]!);
				expect(pptx, SLOTS[slot]).toBeDefined();
				const { chartExSpace, issues } = parseChartExSpace(xml);
				expect(issues).toEqual([]);
				expect(projectNeutralChartEx(chartExSpace), xml.slice(0, 400)).toEqual(
					projectPptxChartEx(pptx!),
				);
			});
		},
	);

	it('retains authored hidden-only labels, content switches and numeric scaling in both public models', async () => {
		const byPart = await chartsByPart(await deckWith([FULL], SLOTS));
		const host = projectPptxChartEx(byPart.get(SLOTS[0]!)!);
		const neutral = projectNeutralChartEx(parseChartExSpace(FULL).chartExSpace);
		for (const chart of [host, neutral]) {
			const series = chart.series[0]!;
			expect(series['labelOptions']).toMatchObject({
				showValue: true,
				showCategoryName: true,
				showSeriesName: false,
			});
			const labels = series['labels'] as Record<string, unknown>[];
			expect(labels.find((label) => label['index'] === 0)).toMatchObject({
				deleted: true,
				showValue: true,
				showCategoryName: true,
				showSeriesName: false,
			});
			expect(labels.find((label) => label['index'] === 2)).toMatchObject({
				showValue: false,
				showCategoryName: false,
				showSeriesName: true,
			});
			expect(labels.find((label) => label['index'] === 2)?.['deleted']).toBeUndefined();
			const axis = chart.axes.find((axis) => axis['id'] === 1)!;
			expect(axis).toMatchObject({ max: 100, majorUnit: 20 });
			expect(axis['min']).toBeUndefined();
			expect(axis['minorUnit']).toBeUndefined();
		}
	});

	it('covers every committed and synthetic part', () => {
		expect(PARTS).toHaveLength(15 + 1 + EDGE.length);
		expect(SLOTS).toHaveLength(7);
	});
});
