import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../index';
import type { ChartPptxElement, PptxData } from '../../index';
import { findInSlides, replaceInSlides } from '../../editor/render/find-replace';
import { createChartLabelDeck } from '../helpers/chart-label-deck';

function chart(data: PptxData): ChartPptxElement {
	return data.slides[0]!.elements.find((element) => element.type === 'chart') as ChartPptxElement;
}

describe('chart label workbook transactions', () => {
	it('searches labels only when enabled and replaces separate categories immutably', async () => {
		const data = await new PptxHandler().load(await createChartLabelDeck());
		expect(findInSlides(data.slides, 'Dataset')).toEqual([]);
		const match = findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]!;
		expect(match.chartSeriesIndex).toBe(0);
		expect(match.readOnly).toBe(false);
		const result = replaceInSlides(data.slides, 'Q', 'Quarter ', { includeChartData: true });
		expect(result.replacements).toBe(2);
		expect(chart({ ...data, slides: [...result.slides] }).chartData!.categories).toEqual([
			'Quarter 1',
			'Quarter 2',
		]);
		expect(chart(data).chartData!.categories).toEqual(['Q1', 'Q2']);
	});
	it('replaces and round-trips a grouped chart while preserving its sibling', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createChartLabelDeck());
		const original = chart(data);
		const sibling = {
			type: 'text' as const,
			id: 'sibling-text',
			x: 20,
			y: 520,
			width: 200,
			height: 40,
			text: 'Untouched sibling',
			textSegments: [{ text: 'Untouched sibling', style: { fontSize: 18 } }],
		};
		data.slides[0]!.elements = [
			{
				type: 'group',
				id: 'chart-group',
				x: 0,
				y: 0,
				width: 900,
				height: 600,
				children: [original, sibling],
			},
		];
		const match = findInSlides(data.slides, 'Q1', { includeChartData: true })[0]!;
		expect(match.elementId).toBe('chart-group');
		expect(match.elementPath).toEqual([0]);
		const result = replaceInSlides(data.slides, 'Q1', 'First quarter', { includeChartData: true });
		const group = result.slides[0]!.elements[0]!;
		if (group.type !== 'group') throw new Error('Expected group');
		expect(group.children[1]).toBe(sibling);
		expect(original.chartData!.categories[0]).toBe('Q1');
		const bytes = await handler.save([...result.slides]);
		const reopened = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		const reopenedGroup = reopened.slides[0]!.elements[0]!;
		if (reopenedGroup.type !== 'group') throw new Error('Expected saved group');
		const savedChart = reopenedGroup.children.find(
			(element) => element.type === 'chart',
		) as ChartPptxElement;
		expect(savedChart.chartData!.categories[0]).toBe('First quarter');
		expect(
			reopenedGroup.children.some(
				(element) => 'text' in element && element.text === 'Untouched sibling',
			),
		).toBe(true);
		handler.dispose();
	});
	it('saves series names and category labels to caches and workbook cells together', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createChartLabelDeck());
		const model = chart(data).chartData!;
		expect(model.dataLabelTextEditable).toBe(true);
		model.series[0]!.name = 'Renamed dataset';
		model.categories[0] = 'First quarter';
		model.dataLabelTextOnly = true;
		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		const workbook = await JSZip.loadAsync(
			await zip.file('ppt/embeddings/Chart_Data.xlsx')!.async('uint8array'),
		);
		const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(sheet).toContain('Renamed dataset');
		expect(sheet).toContain('First quarter');
		expect(sheet).toContain('001 untouched ');
		expect(sheet).toContain('s="7"');
		expect(await workbook.file('unrelated.xml')!.async('string')).toBe(
			'<retained value="opaque"/>',
		);
		const reopened = await new PptxHandler().load(
			bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
		);
		expect(chart(reopened).chartData!.series[0]!.name).toBe('Renamed dataset');
		expect(chart(reopened).chartData!.categories).toEqual(['First quarter', 'Q2']);
		expect(chart(reopened).chartData!.series[0]!.values).toEqual([10, 20]);
	});
	it.each([
		{ formula: true },
		{ external: true },
		{ missing: true },
		{ inconsistent: true },
		{ shared: true },
	])('marks unsupported source %j read-only', async (options) => {
		const data = await new PptxHandler().load(await createChartLabelDeck(options));
		expect(chart(data).chartData!.dataLabelTextEditable).toBe(false);
		expect(findInSlides(data.slides, 'Dataset', { includeChartData: true })[0]!.readOnly).toBe(
			true,
		);
		const original = JSON.stringify(data.slides);
		expect(() =>
			replaceInSlides(data.slides, 'Dataset', 'Changed', { includeChartData: true }),
		).toThrow('source workbook');
		expect(JSON.stringify(data.slides)).toBe(original);
	});
	it('fails a forced transaction instead of returning a cache-only export', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createChartLabelDeck({ formula: true }));
		const model = chart(data).chartData!;
		model.series[0]!.name = 'Renamed dataset';
		model.dataLabelTextOnly = true;
		await expect(handler.save(data.slides)).rejects.toThrow('formula-dependent');
	});
});
