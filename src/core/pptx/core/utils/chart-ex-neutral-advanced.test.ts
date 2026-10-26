import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import {
	createChartExSpace,
	writeChartExSpace,
	parseChartExSpace,
	chartExDataChanged,
} from '../../../chart';
import { PptxHandler } from '../PptxHandler';
import type { PptxChartData, ChartPptxElement } from '../types';
import { saveChartExPartThroughNeutral } from './chart-ex-neutral-save';
import { saveChartExPackageThroughNeutral } from './chart-ex-neutral-package-save';
const before: PptxChartData = {
	chartType: 'waterfall',
	title: 'Title',
	categories: ['A', 'B'],
	series: [{ name: 'Sales', values: [1, 2] }],
	style: { hasTitle: true, hasLegend: true, legendPosition: 'r' },
	axes: [
		{ axisType: 'catAx', axisId: 0 },
		{ axisType: 'valAx', axisId: 1 },
	],
};
function source() {
	return createChartExSpace({ categories: before.categories, series: before.series }, 'waterfall', {
		title: 'Title',
		legend: { visible: true, position: 'r' },
	});
}
it('saves native axis, point labels and DrawingML style changes through shared operations', () => {
	const xml = writeChartExSpace(source()),
		after = structuredClone(before);
	Object.assign(after.style!, {
		chartAreaFill: '#112233',
		titleFontBold: true,
		legendTextStyle: { fontSize: 14, fontFamily: 'Aptos' },
	});
	Object.assign(after.axes![1]!, {
		min: 0,
		max: 10,
		majorUnit: 2,
		minorUnit: 1,
		numFmt: { formatCode: '0.0', sourceLinked: false },
		titleText: 'Axis',
		fontSize: 15,
		tickLblPos: 'none',
		spPr: { strokeColor: '#334455', strokeWidth: 2 },
	});
	after.elementFormats = {
		'series:0': {
			fill: {
				kind: 'pattern',
				pattern: { preset: 'diagCross', foreground: '#123456', background: '#FFFFFF' },
			},
			effects: { softEdge: 2 },
			line: { kind: 'solid', color: '#112233', width: 2, dash: 'dash' },
		},
	};
	after.series[0]!.dataLabelOptions = {
		showValue: false,
		showCategory: true,
		numberFormat: '0',
		txPr: { fontSize: 11 },
	};
	after.series[0]!.dataLabels = [
		{
			idx: 0,
			showVal: true,
			deleted: true,
			position: 'outEnd',
			numberFormat: '0.0',
			txPr: { fontSize: 12, color: '#445566' },
			spPr: { fillColor: '#EEAAAA' },
		},
	];
	const saved = saveChartExPartThroughNeutral(xml, before, after);
	expect(saved).toBeDefined();
	const parsed = parseChartExSpace(saved!);
	expect(parsed.issues).toEqual([]);
	const space = parsed.chartExSpace;
	expect(chartExDataChanged(parseChartExSpace(xml).chartExSpace, space)).toBe(false);
	expect(space.plotArea.axes[1]!.scaling).toMatchObject({
		min: 0,
		max: 10,
		majorUnit: 2,
		minorUnit: 1,
	});
	expect(space.plotArea.axes[1]!.tickLabels).toBeUndefined();
	expect(space.plotArea.axes[1]!.title!.tx!.rich!.text).toBe('Axis');
	expect(space.plotArea.series[0]!.dataLabels!.visibility).toMatchObject({
		value: false,
		categoryName: true,
	});
	expect(space.plotArea.series[0]!.dataLabels!.hidden).toEqual([0]);
	expect(space.plotArea.series[0]!.dataLabels!.labels[0]!.txPr!.sourceXml).toContain('sz="1200"');
	expect(space.plotArea.series[0]!.spPr!.fill).toMatchObject({
		kind: 'pattern',
		preset: 'diagCross',
	});
	expect(space.plotArea.series[0]!.spPr!.effectsXml).toContain('softEdge');
	expect(space.legend!.txPr!.sourceXml).toContain('Aptos');
});
it.each(['logScale', 'crossesAt'] as const)(
	'refuses unsupported %s rather than silently ignoring it',
	async (key) => {
		const after = structuredClone(before);
		Object.assign(after.axes![1]!, { [key]: key === 'logScale' ? true : 2 });
		await expect(
			saveChartExPackageThroughNeutral(
				new JSZip(),
				'ppt/chart.xml',
				writeChartExSpace(source()),
				before,
				after,
			),
		).rejects.toThrow('does not support');
	},
);
it('refuses custom point text which ChartEx has no node to store', async () => {
	const after = structuredClone(before);
	after.series[0]!.dataLabels = [{ idx: 0, text: 'Cannot store this' }];
	await expect(
		saveChartExPackageThroughNeutral(
			new JSZip(),
			'ppt/chart.xml',
			writeChartExSpace(source()),
			before,
			after,
		),
	).rejects.toThrow('point-label');
});
it('embeds a picture chosen for a ChartEx format target using the existing host relationship allocator', async () => {
	const bytes = readFileSync(
		path.join(import.meta.dirname, '../../__tests__/fixtures/e2e/chart-gallery.pptx'),
	);
	const handler = new PptxHandler();
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
	const chart = element.chartData!;
	chart.elementFormats = {
		...chart.elementFormats,
		chartArea: {
			fill: {
				kind: 'picture',
				picture: {
					src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==',
				},
			},
		},
	};
	slide.isDirty = true;
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const xml = await zip.file(chart.chartPartPath!)!.async('string');
	expect(xml).toContain('blipFill');
	expect(xml).toMatch(/r:embed="rId\d+"/u);
	const part = chart.chartPartPath!,
		slash = part.lastIndexOf('/');
	const rels = await zip
		.file(`${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`)!
		.async('string');
	expect(rels).toContain('/image');
	expect(
		Object.keys(zip.files).some((name) => name.startsWith('ppt/media/') && name.endsWith('.png')),
	).toBe(true);
});
