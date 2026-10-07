import { readFileSync } from 'node:fs';

import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { ChartPptxElement, PptxChartType, XmlObject } from '../../types';

const parser = new XMLParser({ ignoreAttributes: false });
const chartData = {
	categories: ['A', 'B', 'C'],
	series: [{ name: 'Series A', values: [0, 2, 3] }],
};

async function generatedDeck(type: PptxChartType): Promise<Uint8Array> {
	const { handler, data, createSlide } = await PptxHandler.createBlank();
	data.slides.push(
		createSlide().addChart(type, chartData, { x: 60, y: 80, width: 700, height: 400 }).build(),
	);
	return handler.save(data.slides);
}

function importedChartXml(tag: string, hasSeriesAxis: boolean): string {
	const axis = (kind: string, id: number, cross: number, pos: string) =>
		`<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/></c:${kind}>`;
	return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
		<c:chart><c:plotArea><c:layout/><c:${tag}>
		${tag === 'line3DChart' ? '<c:grouping val="standard"/>' : '<c:wireframe val="0"/>'}
		<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series A</c:v></c:tx>
		<c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat>
		<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>0</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser>
		<c:axId val="41"/><c:axId val="82"/>${hasSeriesAxis ? '<c:axId val="123"/>' : ''}
		</c:${tag}>${axis('catAx', 41, 82, 'b')}${axis('valAx', 82, 41, 'l')}${hasSeriesAxis ? axis('serAx', 123, 82, 'b') : ''}
		</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

async function chartPlot(bytes: Uint8Array, part = 'ppt/charts/chart1.xml'): Promise<XmlObject> {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file(part)!.async('string');
	return parser.parse(xml)['c:chartSpace']['c:chart']['c:plotArea'];
}

function nodes(value: unknown): XmlObject[] {
	return value === undefined ? [] : Array.isArray(value) ? value : [value as XmlObject];
}

function assertAxes(plot: XmlObject, tag: string, ids: number[]): void {
	const references = nodes((plot[`c:${tag}`] as XmlObject)['c:axId']).map((n) =>
		Number(n['@_val']),
	);
	expect(references).toEqual(ids);
	const definitions = ['catAx', 'valAx', 'dateAx', 'serAx'].flatMap((name) =>
		nodes(plot[`c:${name}`]),
	);
	expect(definitions).toHaveLength(ids.length);
	for (const id of references) {
		expect(
			definitions.filter((n) => Number((n['c:axId'] as XmlObject)['@_val']) === id),
		).toHaveLength(1);
	}
	for (const axis of definitions) {
		const cross = Number((axis['c:crossAx'] as XmlObject)['@_val']);
		expect(ids).toContain(cross);
	}
}

describe('native chart axes through load/save and editing', () => {
	it('keeps axes and 3-D containers in the PowerPoint COM-authored fixture after editing', async () => {
		const source = readFileSync(
			new URL(
				'../../../__tests__/fixtures/e2e/three-d-parity/three-d-charts.pptx',
				import.meta.url,
			),
		);
		const handler = new PptxHandler();
		const loaded = await handler.load(
			source.buffer.slice(source.byteOffset, source.byteOffset + source.byteLength),
		);
		const expected: Array<{ part: string; tag: string; ids: number[] }> = [];
		for (const slide of loaded.slides) {
			for (const element of slide.elements) {
				if (element.type !== 'chart' || !element.chartData) {
					continue;
				}
				const chart = element.chartData;
				if (chart.chartType !== 'line3D' && chart.chartType !== 'surface') {
					continue;
				}
				const tag = chart.chartType === 'line3D' ? 'line3DChart' : 'surface3DChart';
				const part = chart.chartPartPath!;
				const plot = await chartPlot(source, part);
				const ids = nodes((plot[`c:${tag}`] as XmlObject)['c:axId']).map((node) =>
					Number(node['@_val']),
				);
				expect(ids).toHaveLength(3);
				expected.push({ part, tag, ids });
				chart.series[0].values[0] = 0;
				slide.isDirty = true;
			}
		}
		expect(expected).toHaveLength(3);
		const saved = await handler.save(loaded.slides);
		for (const { part, tag, ids } of expected) {
			assertAxes(await chartPlot(saved, part), tag, ids);
		}
		const reloaded = await new PptxHandler().load(saved.buffer as ArrayBuffer);
		const charts = reloaded.slides
			.flatMap((slide) => slide.elements)
			.filter(
				(element): element is ChartPptxElement =>
					element.type === 'chart' &&
					expected.some((entry) => entry.part === element.chartData?.chartPartPath),
			);
		expect(charts).toHaveLength(3);
		for (const chart of charts) {
			expect(chart.chartData!.series[0].values[0]).toBe(0);
			if (chart.chartData!.chartType === 'surface') {
				expect(chart.chartData!.surfaceTopView).toBe(false);
			}
		}
	});

	it.each(['line3D', 'surface'] as const)(
		'keeps generated %s axes and numeric zero on repeated saves',
		async (type) => {
			let bytes = await generatedDeck(type);
			for (let pass = 0; pass < 2; pass++) {
				const handler = new PptxHandler();
				const loaded = await handler.load(bytes.buffer as ArrayBuffer);
				const chart = loaded.slides[0].elements.find((e) => e.type === 'chart') as ChartPptxElement;
				expect(chart.chartData!.chartType).toBe(type);
				expect(chart.chartData!.series[0].values).toEqual([0, 2, 3]);
				expect(chart.chartData!.axes).toHaveLength(3);
				bytes = await handler.save(loaded.slides);
				assertAxes(
					await chartPlot(bytes),
					type === 'line3D' ? 'line3DChart' : 'surfaceChart',
					[111111111, 222222222, 333333333],
				);
			}
		},
	);

	it.each([
		['line3DChart', true],
		['surface3DChart', true],
		['surfaceChart', true],
		['surfaceChart', false],
	] as const)(
		'preserves imported %s (series axis: %s) and its original IDs',
		async (tag, hasSeriesAxis) => {
			// Use independent ChartML with Office-style IDs, not the generator under test.
			const zip = await JSZip.loadAsync(await generatedDeck('line'));
			zip.file('ppt/charts/chart1.xml', importedChartXml(tag, hasSeriesAxis));
			const bytes = await zip.generateAsync({ type: 'uint8array' });
			const handler = new PptxHandler();
			const loaded = await handler.load(bytes.buffer as ArrayBuffer);
			const chart = loaded.slides[0].elements.find((e) => e.type === 'chart') as ChartPptxElement;
			const ids = hasSeriesAxis ? [41, 82, 123] : [41, 82];
			expect(chart.chartData!.axes!.map((axis) => axis.axisId)).toEqual(ids);
			expect(chart.chartData!.series[0].values).toEqual([0, 2, 3]);
			if (tag.startsWith('surface')) {
				expect(chart.chartData!.surfaceTopView).toBe(tag === 'surfaceChart');
			}
			const firstSaved = await handler.save(loaded.slides);
			assertAxes(await chartPlot(firstSaved), tag, ids);
			// An edit must not discard the series axis or its crossing relationship.
			chart.chartData!.series[0].values[1] = 7;
			if (hasSeriesAxis) {
				chart.chartData!.axes!.find((axis) => axis.axisType === 'serAx')!.tickLabelSkip = 2;
			}
			loaded.slides[0].isDirty = true;
			const edited = await handler.save(loaded.slides);
			const plot = await chartPlot(edited);
			assertAxes(plot, tag, ids);
			if (hasSeriesAxis) {
				expect((plot['c:serAx'] as XmlObject)['c:tickLblSkip']).toEqual({ '@_val': '2' });
				expect((plot['c:serAx'] as XmlObject)['c:crossAx']).toEqual({ '@_val': '82' });
			}
			const reloaded = await new PptxHandler().load(edited.buffer as ArrayBuffer);
			const result = reloaded.slides[0].elements.find(
				(e) => e.type === 'chart',
			) as ChartPptxElement;
			expect(result.chartData!.series[0].values).toEqual([0, 7, 3]);
		},
	);
});
