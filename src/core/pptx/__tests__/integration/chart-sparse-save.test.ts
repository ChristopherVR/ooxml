import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import type { ChartPptxElement, PptxData } from '../../index';
import { requireFixture } from '../require-fixture';

/**
 * A chart range is often mostly empty: Excel caches only the cells that hold values and declares
 * how many the range has. Saving such a chart must keep that count and those points, on both of
 * PowerPoint's save paths, and must not write its blanks into the embedded workbook.
 */
const FIXTURE = requireFixture(
	fileURLToPath(new URL('../fixtures/e2e/chart-data-fidelity.pptx', import.meta.url)),
);
const EMBEDDING = 'ppt/embeddings/Microsoft_Excel_Worksheet2.xlsx';
const HUGE = 2_000_000_000;

const buffer = (bytes: Uint8Array) =>
	bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
const chartOf = (data: PptxData) =>
	data.slides[2]!.elements.find(
		(element): element is ChartPptxElement => element.type === 'chart',
	)!;

/** The fixture with the first series' value cache replaced; `unmodelled` forces the tree writer. */
async function deck(cache: string, unmodelled: boolean) {
	const source = readFileSync(FIXTURE);
	const zip = await JSZip.loadAsync(source);
	const path = chartOf(await new PptxHandler().load(buffer(source))).chartData!.chartPartPath!;
	let xml = await zip.file(path)!.async('string');
	xml = xml.replace(
		/(<c:val>[\s\S]*?<c:numCache>)[\s\S]*?(<\/c:numCache>)/u,
		(_, open: string, close: string) =>
			`${open}<c:formatCode>General</c:formatCode>${cache}${close}`,
	);
	if (unmodelled) xml = xml.replace('</c:ser>', '<c:notInTheModel/></c:ser>');
	zip.file(path, xml);
	return { path, bytes: await zip.generateAsync({ type: 'uint8array' }) };
}

/** The value cache of the first series in a saved part. */
const valueCache = (xml: string) =>
	/<c:val>[\s\S]*?<c:numCache>([\s\S]*?)<\/c:numCache>/u.exec(xml)![1]!;
const pointIndices = (cache: string) =>
	[...cache.matchAll(/<c:pt idx="(\d+)"/gu)].map((match) => Number(match[1]));
const declaredCount = (cache: string) => Number(/<c:ptCount val="(\d+)"/u.exec(cache)![1]);

async function saveEdited(
	source: { path: string; bytes: Uint8Array },
	edit: (chart: NonNullable<ChartPptxElement['chartData']>) => void,
) {
	const handler = new PptxHandler();
	const data = await handler.load(buffer(source.bytes));
	edit(chartOf(data).chartData!);
	const saved = await JSZip.loadAsync(await handler.save(data.slides));
	return { saved, cache: valueCache(await saved.file(source.path)!.async('string')) };
}

const SPARSE =
	'<c:ptCount val="5"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="3"><c:v>31</c:v></c:pt>';

describe.each([
	['the shared chart writer', false],
	['the object-tree writer', true],
])('saving a sparse chart through %s', (_, unmodelled) => {
	it('keeps the count and points through a title edit', async () => {
		const { cache } = await saveEdited(await deck(SPARSE, unmodelled), (chart) => {
			chart.title = 'Edited';
		});
		expect(declaredCount(cache)).toBe(5);
		expect(pointIndices(cache)).toEqual([0, 3]);
	});

	it('changes only the edited point and invents none', async () => {
		const { cache } = await saveEdited(await deck(SPARSE, unmodelled), (chart) => {
			chart.series[0]!.values[3] = 99;
		});
		expect(declaredCount(cache)).toBe(5);
		expect(pointIndices(cache)).toEqual([0, 3]);
		expect(cache).toContain('<c:pt idx="3"><c:v>99</c:v>');
	});

	it('keeps a declared count too large to show, and the point past the view', async () => {
		const far = `<c:ptCount val="${HUGE}"/><c:pt idx="0"><c:v>12</c:v></c:pt><c:pt idx="${HUGE - 1}"><c:v>7</c:v></c:pt>`;
		const { cache } = await saveEdited(await deck(far, unmodelled), (chart) => {
			expect(chart.series[0]!.values).toHaveLength(2);
			chart.series[0]!.values[0] = 99;
		});
		expect(declaredCount(cache)).toBe(HUGE);
		expect(pointIndices(cache)).toEqual([0, HUGE - 1]);
		expect(cache).toContain('<c:pt idx="0"><c:v>99</c:v>');
	});
});

describe('the embedded workbook of a sparse chart', () => {
	const cells = async (saved: JSZip) => {
		const book = await JSZip.loadAsync(await saved.file(EMBEDDING)!.async('uint8array'));
		const sheet = await book.file('xl/worksheets/sheet1.xml')!.async('string');
		return (ref: string) => new RegExp(`<c r="${ref}"[^>]*><v>([^<]*)</v>`, 'u').exec(sheet)?.[1];
	};

	it.each([
		['the shared chart writer', false],
		['the object-tree writer', true],
	])('writes points and leaves blank cells alone through %s', async (_, unmodelled) => {
		const { saved } = await saveEdited(await deck(SPARSE, unmodelled), (chart) => {
			chart.series[0]!.values[3] = 99;
		});
		const cell = await cells(saved);
		expect(cell('B5')).toBe('99');
		// B3 and B4 hold values the cache does not show, and B7 is past the range.
		expect(cell('B3')).toBe('25');
		expect(cell('B4')).toBe('7');
		expect(cell('B7')).toBeUndefined();
	});

	it('empties the cell of a point the edit removed', async () => {
		const { saved, cache } = await saveEdited(await deck(SPARSE, false), (chart) => {
			chart.series[0]!.blanks![3] = true;
		});
		expect(pointIndices(cache)).toEqual([0]);
		expect((await cells(saved))('B5')).toBeUndefined();
	});
});
