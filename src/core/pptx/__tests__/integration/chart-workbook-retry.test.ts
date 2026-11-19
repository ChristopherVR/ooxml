import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';

import { getCell } from '../../../xlsx/cells';
import { loadXlsx } from '../../../xlsx/read/index';
import * as workbookSave from '../../core/core/runtime/chart-external-workbook-save';
import { PptxHandler } from '../../index';
import type { ChartPptxElement, PptxSlide } from '../../index';
import { requireFixture } from '../require-fixture';

// A deck PowerPoint saved with a filtered row: each value cache holds 3 points of a 4-row range, so
// value writes cannot be placed and every save that changes a value lands only part of its writes.
const FIXTURE = requireFixture(
	fileURLToPath(new URL('../fixtures/e2e/chart-filtered-series.pptx', import.meta.url)),
);
const CHART = 'ppt/charts/chart1.xml';
const WORKBOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx';

async function deck(edit?: (chartXml: string) => string): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
	if (edit) zip.file(CHART, edit(await zip.file(CHART)!.async('string')));
	return zip.generateAsync({ type: 'arraybuffer' });
}

/** The deck with a second chart, a copy of the first, reading the same embedded workbook. */
async function sharedDeck(): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
	const text = async (name: string) => zip.file(name)!.async('string');
	zip.file('ppt/charts/chart2.xml', await text(CHART));
	zip.file('ppt/charts/_rels/chart2.xml.rels', await text('ppt/charts/_rels/chart1.xml.rels'));
	const slide = await text('ppt/slides/slide1.xml');
	const frame = /<p:graphicFrame>.*<\/p:graphicFrame>/su.exec(slide)![0];
	const copy = frame
		.replace('id="2" name="Chart 1"', 'id="3" name="Chart 2"')
		.replace('r:id="rId2"', 'r:id="rId3"')
		.replace(/\{1E9F4A36[^}]*\}/u, '{1E9F4A36-8B0D-12EC-6444-50CB5B6E3EE5}');
	zip.file('ppt/slides/slide1.xml', slide.replace(frame, frame + copy));
	zip.file(
		'ppt/slides/_rels/slide1.xml.rels',
		(await text('ppt/slides/_rels/slide1.xml.rels')).replace(
			'</Relationships>',
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>',
		),
	);
	zip.file(
		'[Content_Types].xml',
		(await text('[Content_Types].xml')).replace(
			'</Types>',
			'<Override PartName="/ppt/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
		),
	);
	return zip.generateAsync({ type: 'arraybuffer' });
}

const chartsOf = (slides: PptxSlide[]) =>
	slides
		.flatMap((slide) => slide.elements)
		.filter((element): element is ChartPptxElement => element.type === 'chart')
		.map((element) => element.chartData!);

const chartOf = (slides: PptxSlide[]) =>
	slides
		.flatMap((slide) => slide.elements)
		.find((element): element is ChartPptxElement => element.type === 'chart')!.chartData!;

/** Cells A1:D5 of the saved workbook, as text. */
async function cells(saved: ArrayBuffer): Promise<string[]> {
	const zip = await JSZip.loadAsync(saved);
	const book = await loadXlsx(await zip.file(WORKBOOK)!.async('uint8array'));
	const out: string[] = [];
	for (let row = 0; row < 5; row++)
		for (let col = 0; col < 4; col++)
			out.push(String(getCell(book.sheets[0]!, row, col)?.value ?? ''));
	return out;
}

const B1 = 1;
const D1 = 3;

describe('PowerPoint chart workbook write-back that only partly lands', () => {
	it('puts back a renamed series when the edit is undone', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await deck());
		const loaded = await cells(await handler.save(structuredClone(data.slides)));
		const original = structuredClone(data.slides);
		const chart = chartOf(data.slides);
		chart.series[0]!.name = 'Renamed';
		chart.series[0]!.values[0] = 4242;
		expect((await cells(await handler.save(data.slides)))[B1]).toBe('Renamed');
		const undone = await handler.save(original);
		expect(await cells(undone)).toEqual(loaded);
		expect(await (await JSZip.loadAsync(undone)).file(CHART)!.async('string')).toContain(
			'<c:v>Series A</c:v>',
		);
		handler.dispose();
	});

	it('follows the chart through edits, a failed save, a retry and undo', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await deck());
		const original = structuredClone(data.slides);
		const loaded = await cells(await handler.save(structuredClone(original)));
		const name = async (slides: PptxSlide[]) => (await cells(await handler.save(slides)))[B1];

		const first = structuredClone(original);
		chartOf(first).series[0]!.name = 'One';
		chartOf(first).series[0]!.values[1] = 7;
		expect(await name(first)).toBe('One');

		const second = structuredClone(first);
		chartOf(second).series[0]!.name = 'Two';
		expect(await name(second)).toBe('Two');

		const third = structuredClone(second);
		chartOf(third).series[0]!.name = 'Three';
		const actual = workbookSave.saveChartExternalWorkbookUpdates;
		const spy = vi
			.spyOn(workbookSave, 'saveChartExternalWorkbookUpdates')
			.mockResolvedValueOnce({ complete: false })
			.mockImplementation(actual);
		try {
			// The failed save changes nothing; the next one writes the whole edit.
			expect(await name(third)).toBe('Series A');
			expect(await name(third)).toBe('Three');
		} finally {
			spy.mockRestore();
		}

		expect(await cells(await handler.save(structuredClone(original)))).toEqual(loaded);
		expect(await cells(await handler.save(structuredClone(original)))).toEqual(loaded);
		handler.dispose();
	});

	it('keeps what another chart sharing the workbook wrote when it retries', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await sharedDeck());
		const [first, second] = chartsOf(data.slides);
		first!.series[0]!.name = 'A-one';
		first!.series[0]!.values[0] = 4242;
		second!.series[1]!.name = 'C-two';
		const saved = await cells(await handler.save(data.slides));
		expect([saved[B1], saved[D1]]).toEqual(['A-one', 'C-two']);
		const retried = await cells(await handler.save(data.slides));
		expect([retried[B1], retried[D1]]).toEqual(['A-one', 'C-two']);
		handler.dispose();
	});

	it('puts back a renamed series whose sibling names a sheet the workbook lacks', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(
			await deck((xml) => xml.replace('<c:f>Sheet1!$D$1</c:f>', '<c:f>Missing!$D$1</c:f>')),
		);
		const original = structuredClone(data.slides);
		const loaded = await cells(await handler.save(structuredClone(original)));
		const chart = chartOf(data.slides);
		for (const series of chart.series) series.name = `${series.name} renamed`;
		const edited = await cells(await handler.save(data.slides));
		expect(edited[B1]).toBe('Series A renamed');
		expect(edited[D1]).toBe(loaded[D1]);
		expect(await cells(await handler.save(original))).toEqual(loaded);
		handler.dispose();
	});
});
