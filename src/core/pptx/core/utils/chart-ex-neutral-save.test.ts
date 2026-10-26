import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { PptxHandler } from '../PptxHandler';
import type { ChartPptxElement } from '../types';

import { createChartWorkbook, parseChartExSpace, chartExSpaceDataSheet } from '../../../chart';
import { loadXlsx } from '../../../xlsx';
import { saveChartExPackageThroughNeutral } from './chart-ex-neutral-package-save';
import type { PptxChartData } from '../types';
import { planChartExNeutralSave, saveChartExPartThroughNeutral } from './chart-ex-neutral-save';

const PART =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Old</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel" uniqueId="{U}"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Sales</cx:v></cx:txData></cx:tx><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea><cx:legend pos="t" align="ctr" overlay="0" /></cx:chart></cx:chartSpace>';

const BEFORE: PptxChartData = {
	chartType: 'funnel',
	title: 'Old',
	categories: ['A', 'B'],
	categoryLevels: [['A', 'B']],
	series: [{ name: 'Sales', values: [1, 2] }],
	style: { hasTitle: true, hasLegend: true, legendPosition: 't' },
};

const edited = (recipe: (chart: PptxChartData) => void): PptxChartData => {
	const chart = structuredClone(BEFORE);
	recipe(chart);
	return chart;
};

async function packageFixture() {
	const zip = new JSZip();
	const book = await createChartWorkbook({
		categories: ['A', 'B'],
		series: [{ name: 'Sales', values: [1, 2] }],
	});
	zip.file('ppt/embeddings/data.xlsx', book);
	zip.file(
		'ppt/charts/_rels/chartEx1.xml.rels',
		'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/data.xlsx"/></Relationships>',
	);
	const xml = PART.replace('<cx:chartData>', '<cx:chartData><cx:externalData r:id="rId1"/>')
		.replace('<cx:f>Sheet1!$B$1</cx:f><cx:v>Old</cx:v>', '<cx:v>Old</cx:v>')
		.replace('layoutId="funnel"', 'layoutId="boxWhisker"');
	const before = { ...structuredClone(BEFORE), chartType: 'boxWhisker' as const };
	return { zip, book, xml, before, part: 'ppt/charts/chartEx1.xml' };
}

it('synchronizes cache values and structural growth with the embedded workbook', async () => {
	const { zip, book, xml, before, part } = await packageFixture();
	const after = structuredClone(before);
	after.series.push({ name: 'Second', values: [7, 8] });
	const expanded = await saveChartExPackageThroughNeutral(zip, part, xml, before, after);
	const bytes = await zip.file('ppt/embeddings/data.xlsx')!.async('uint8array');
	const original = await JSZip.loadAsync(book),
		result = await JSZip.loadAsync(bytes);
	expect(await result.file('xl/worksheets/sheet1.xml')!.async('string')).toBe(
		await original.file('xl/worksheets/sheet1.xml')!.async('string'),
	);
	expect((await loadXlsx(bytes)).sheets).toHaveLength(2);
	expect(
		chartExSpaceDataSheet(parseChartExSpace(expanded!).chartExSpace).series[1]!.values,
	).toEqual([7, 8]);
	expect(expanded).toContain('ChartExData');
});

it('refuses external data edits atomically while allowing a layout-only change', async () => {
	const { zip, xml, before, part, book } = await packageFixture();
	const rels = 'ppt/charts/_rels/chartEx1.xml.rels';
	zip.file(
		rels,
		(await zip.file(rels)!.async('string')).replace(
			'Target="../embeddings/data.xlsx"',
			'TargetMode="External" Target="https://example.test/data.xlsx"',
		),
	);
	const after = structuredClone(before);
	after.series[0]!.values[0] = 4;
	await expect(saveChartExPackageThroughNeutral(zip, part, xml, before, after)).rejects.toThrow(
		'local embedded',
	);
	expect(await zip.file('ppt/embeddings/data.xlsx')!.async('uint8array')).toEqual(book);
	const layout = { ...structuredClone(before), chartType: 'treemap' as const };
	expect(await saveChartExPackageThroughNeutral(zip, part, xml, before, layout)).toContain(
		'layoutId="treemap"',
	);
});

it('reports original gallery orphan sources to the PowerPoint caller', async () => {
	const handler = new PptxHandler();
	const bytes = readFileSync(
		path.join(import.meta.dirname, '../../__tests__/fixtures/e2e/chart-gallery.pptx'),
	);
	const data = await handler.load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	const slide = data.slides.find((slide) =>
		slide.elements.some(
			(element) => (element as ChartPptxElement).chartData?.chartType === 'waterfall',
		),
	)!;
	const element = slide.elements.find(
		(element) => (element as ChartPptxElement).chartData?.chartType === 'waterfall',
	) as ChartPptxElement;
	element.chartData!.series[0]!.values[0] = 123;
	slide.isDirty = true;
	await expect(handler.save(data.slides)).rejects.toThrow('no embedded workbook');
});

describe('saveChartExPartThroughNeutral', () => {
	it('claims explicit label content accompanying a show-labels action', () => {
		const after = edited((chart) => {
			chart.style!.hasDataLabels = true;
			chart.style!.dataLabels = { showValue: true, showCategory: false };
		});
		const xml = saveChartExPartThroughNeutral(PART, BEFORE, after)!;
		const labels = parseChartExSpace(xml).chartExSpace.plotArea.series[0]!.dataLabels!;
		expect(labels.visibility).toMatchObject({ value: true, categoryName: false });
		expect(xml).not.toContain('formatCode="General"');
	});
	it('saves claimed edits and keeps every other byte (formulas, ids, empty-element form)', () => {
		const after = edited((chart) => {
			chart.series[0]!.name = 'Revenue';
			chart.series[0]!.values = [5, 6];
			chart.style!.hasLegend = false;
		});
		expect(saveChartExPartThroughNeutral(PART, BEFORE, after)).toBe(
			PART.replace('<cx:v>Sales</cx:v>', '<cx:v>Revenue</cx:v>')
				.replace('<cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt>', '')
				.replace(
					'<cx:lvl ptCount="2"></cx:lvl></cx:numDim>',
					'<cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">5</cx:pt><cx:pt idx="1">6</cx:pt></cx:lvl></cx:numDim>',
				)
				.replace('<cx:legend pos="t" align="ctr" overlay="0" />', ''),
		);
	});

	it('writes a new title over a cell-linked one and removes a hidden one', () => {
		const titled = saveChartExPartThroughNeutral(
			PART,
			BEFORE,
			edited((chart) => {
				chart.title = 'New';
			}),
		);
		expect(titled).toContain(
			'<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr /><a:p><a:r><a:t>New</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>',
		);
		const hidden = saveChartExPartThroughNeutral(
			PART,
			BEFORE,
			edited((chart) => {
				chart.style!.hasTitle = false;
			}),
		);
		expect(hidden).not.toContain('<cx:title');
	});

	it('leaves unclaimed edits and unreadable parts to the object-tree update', () => {
		const coloured = edited((chart) => {
			chart.series[0]!.color = '#FF0000';
		});
		expect(planChartExNeutralSave(BEFORE, coloured)).toBeDefined();
		expect(saveChartExPartThroughNeutral(PART, BEFORE, coloured)).toContain('FF0000');
		const added = edited((chart) => {
			chart.series.push({ name: 'More', values: [1] });
		});
		expect(planChartExNeutralSave(BEFORE, added)).toBeUndefined();
		const renamed = edited((chart) => {
			chart.series[0]!.name = 'X';
		});
		expect(
			saveChartExPartThroughNeutral(
				PART.replace('</cx:series>', '<cx:x/></cx:series>'),
				BEFORE,
				renamed,
			),
		).toBeUndefined();
		expect(planChartExNeutralSave(BEFORE, structuredClone(BEFORE))).toEqual([]);
	});
	it('switches supported families without regenerating producer metadata or hierarchy', () => {
		const xml = PART.replace('layoutId="funnel"', 'layoutId="treemap"')
			.replace('type="val"', 'type="size"')
			.replace(
				'</cx:strDim>',
				'<cx:lvl ptCount="2"><cx:pt idx="0">Parent</cx:pt><cx:pt idx="1">Parent</cx:pt></cx:lvl></cx:strDim>',
			);
		const before = {
			...structuredClone(BEFORE),
			chartType: 'treemap' as const,
			categoryLevels: [
				['A', 'B'],
				['Parent', 'Parent'],
			],
		};
		const after = {
			...structuredClone(before),
			chartType: 'sunburst' as const,
			categories: ['Leaf', 'B'],
		};
		const written = saveChartExPartThroughNeutral(xml, before, after)!;
		expect(written).toContain('uniqueId="{U}"');
		const space = parseChartExSpace(written).chartExSpace;
		expect(space.plotArea.series[0]!.layoutId).toBe('sunburst');
		expect(space.chartData.data[0]!.dimensions[0]!.levels[1]!.points[0]!.value).toBe('Parent');
		expect(space.chartData.data[0]!.dimensions[0]!.levels[0]!.points[0]!.value).toBe('Leaf');
	});
	it('adds independent box-whisker series and applies family options through shared operations', () => {
		const before = { ...structuredClone(BEFORE), chartType: 'boxWhisker' as const };
		const after = structuredClone(before);
		after.series[0]!.boxWhiskerOptions = { showMeanLine: true, quartileMethod: 'inclusive' };
		after.series.push({ name: 'New', values: [7, 8], color: '#123456' });
		const xml = saveChartExPartThroughNeutral(
			PART.replace('layoutId="funnel"', 'layoutId="boxWhisker"'),
			before,
			after,
		)!;
		const space = parseChartExSpace(xml).chartExSpace;
		expect(space.plotArea.series).toHaveLength(2);
		expect(space.chartData.data).toHaveLength(2);
		expect(space.plotArea.series[0]!.layoutProperties!.visibility!.meanLine).toBe(true);
		expect(space.chartData.data[1]!.dimensions[1]!.levels[0]!.points.map((p) => p.value)).toEqual([
			'7',
			'8',
		]);
		expect(space.plotArea.series[1]!.spPr!.fill).toMatchObject({
			kind: 'solid',
			color: { value: '123456' },
		});
	});
});
