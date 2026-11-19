import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { getCell } from '../../../xlsx/cells';
import { loadXlsx } from '../../../xlsx/read/index';
import { ChartPartRevisions } from '../../core/core/runtime/chart-part-revisions';
import { chartDataApplySheetEdits } from '../../core/utils/chart-data-sheet';
import { cloneSlide } from '../../core/utils/clone-utils';
import { PptxHandler } from '../../index';
import type { ChartPptxElement, PptxSlide } from '../../index';
import { requireFixture } from '../require-fixture';

// A PowerPoint-authored column chart reading Sheet1!$A$2:$B$6 of its embedded workbook.
const FIXTURE = requireFixture(
	fileURLToPath(new URL('../fixtures/e2e/chart-data-fidelity.pptx', import.meta.url)),
);
const CHART = 'ppt/charts/chart3.xml';
const WORKBOOK = 'ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx';

const buffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const chartOf = (slides: PptxSlide[]) =>
	slides[2]!.elements.find((element): element is ChartPptxElement => element.type === 'chart')!;
const removeRow = (slides: PptxSlide[]) => {
	const chart = chartOf(slides);
	chart.chartData = chartDataApplySheetEdits(chart.chartData!, [
		{ kind: 'remove-category', category: 1 },
	])!;
};

/** The formulas and cached values of the saved chart part, and cells A1:B7 of its workbook. */
async function saved(bytes: ArrayBuffer) {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file(CHART)!.async('string');
	const book = await loadXlsx(await zip.file(WORKBOOK)!.async('uint8array'));
	const cells: string[] = [];
	for (let row = 0; row < 7; row++)
		for (let col = 0; col < 2; col++)
			cells.push(String(getCell(book.sheets[0]!, row, col)?.value ?? ''));
	return {
		formulas: [...xml.matchAll(/<c:f>([^<]*)<\/c:f>/gu)].map((match) => match[1]),
		values: [...xml.matchAll(/<c:v>([^<]*)<\/c:v>/gu)].map((match) => match[1]),
		cells,
	};
}

/** Each value formula of the saved part read from the workbook, beside the values it caches. */
async function workbookReads(bytes: ArrayBuffer) {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file(CHART)!.async('string');
	const book = await loadXlsx(await zip.file(WORKBOOK)!.async('uint8array'));
	return [...xml.matchAll(/<c:val>[\s\S]*?<\/c:val>/gu)].map(([val]) => {
		const [, sheet, column, first, last] =
			/<c:f>([^!<]+)!\$([A-Z]+)\$(\d+):\$[A-Z]+\$(\d+)<\/c:f>/u.exec(val)!;
		const worksheet = book.sheets.find((entry) => entry.name === sheet)!;
		const col = column!.charCodeAt(0) - 65;
		const read = [];
		for (let row = Number(first) - 1; row < Number(last); row++)
			read.push(String(getCell(worksheet, row, col)?.value ?? ''));
		const cached = [...val.matchAll(/<c:v>([^<]*)<\/c:v>/gu)].map((match) => match[1]);
		return { read, cached };
	});
}

async function open(revisions?: ChartPartRevisions) {
	const handler = new PptxHandler();
	if (revisions)
		(
			handler as unknown as { runtime: { chartPartRevisions: ChartPartRevisions } }
		).runtime.chartPartRevisions = revisions;
	const data = await handler.load(buffer(readFileSync(FIXTURE)));
	return { handler, data };
}

describe('undoing a chart row edit after the chart was saved', () => {
	it('saves the part and workbook the edit started from', async () => {
		const { handler, data } = await open();
		const loaded = await saved(await handler.save(data.slides.map(cloneSlide)));
		const history = data.slides.map(cloneSlide);
		removeRow(data.slides);
		const edited = await saved(await handler.save(data.slides));
		expect(edited.formulas).toContain('Sheet1!$A$2:$A$5');
		expect(await saved(await handler.save(history.map(cloneSlide)))).toEqual(loaded);
		expect(await saved(await handler.save(data.slides))).toEqual(edited);
		handler.dispose();
	});

	it('saves data edited from a released state as edited from the part as it is now', async () => {
		const { handler, data } = await open(new ChartPartRevisions(0));
		const history = data.slides.map(cloneSlide);
		const shown = structuredClone(chartOf(history).chartData!);
		removeRow(data.slides);
		await handler.save(data.slides);
		const undone = await handler.save(history.map(cloneSlide));
		// The part shows the data that was saved, and its workbook holds what the part caches.
		for (const { read, cached } of await workbookReads(undone)) expect(read).toEqual(cached);
		const reloaded = chartOf((await new PptxHandler().load(undone)).slides).chartData!;
		expect(reloaded.categories).toEqual(shown.categories);
		expect(reloaded.series.map((series) => series.values)).toEqual(
			shown.series.map((series) => series.values),
		);
		handler.dispose();
	});

	it('saves a second state from a released one with ranges that read its own data', async () => {
		const { handler, data } = await open(new ChartPartRevisions(0));
		const loaded = data.slides.map(cloneSlide);
		removeRow(data.slides);
		await handler.save(data.slides);
		const chart = chartOf(data.slides);
		chart.chartData = chartDataApplySheetEdits(chart.chartData!, [{ kind: 'add-series' }])!;
		await handler.save(data.slides);
		const twoSeries = data.slides.map(cloneSlide);
		const shown = structuredClone(chartOf(twoSeries).chartData!);
		// Both states are released: the first save moves the loaded data onto a new sheet, and the
		// second saves data whose rows and series that part no longer has.
		await handler.save(loaded.map(cloneSlide));
		const redone = await handler.save(twoSeries.map(cloneSlide));
		const reads = await workbookReads(redone);
		expect(reads).toHaveLength(2);
		for (const { read, cached } of reads) expect(read).toEqual(cached);
		const reloaded = chartOf((await new PptxHandler().load(redone)).slides).chartData!;
		expect(reloaded.series.map((series) => series.values)).toEqual(
			shown.series.map((series) => series.values),
		);
		handler.dispose();
	});
});

describe('undoing a row edit of a chart over a hidden row', () => {
	// PowerPoint saved this chart with a hidden row: each cache holds 3 points of a 4-row range.
	const HIDDEN = requireFixture(
		fileURLToPath(new URL('../fixtures/e2e/chart-filtered-series.pptx', import.meta.url)),
	);

	it('gives back the part as loaded, hidden row included', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(buffer(readFileSync(HIDDEN)));
		const part = async (bytes: ArrayBuffer) =>
			(await JSZip.loadAsync(bytes)).file('ppt/charts/chart1.xml')!.async('string');
		const loaded = await part(await handler.save(data.slides.map(cloneSlide)));
		const history = data.slides.map(cloneSlide);
		const chart = data.slides
			.flatMap((slide) => slide.elements)
			.find((element): element is ChartPptxElement => element.type === 'chart')!;
		chart.chartData = chartDataApplySheetEdits(chart.chartData!, [
			{ kind: 'remove-category', category: 1 },
		])!;
		await handler.save(data.slides);
		expect(await part(await handler.save(history.map(cloneSlide)))).toBe(loaded);
		handler.dispose();
	});
});
