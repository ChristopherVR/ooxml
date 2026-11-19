import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PptxHandler } from '../../index';
import type { ChartPptxElement } from '../../index';
import { createChartLabelDeck } from '../helpers/chart-label-deck';

afterEach(() => vi.restoreAllMocks());

/** The label deck with its chart frame shown `frames` times and `extraParts` more rels parts. */
async function deck(frames: number, extraParts: number): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(await createChartLabelDeck());
	const path = 'ppt/slides/slide1.xml';
	const slide = await zip.file(path)!.async('string');
	const frame = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/u.exec(slide)![0];
	zip.file(path, slide.replace(frame, frame.repeat(frames)));
	for (let index = 0; index < extraParts; index++)
		zip.file(
			`ppt/customXml/_rels/item${index}.xml.rels`,
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
		);
	return zip.generateAsync({ type: 'arraybuffer' });
}

async function load(bytes: ArrayBuffer) {
	const clone = vi.spyOn(globalThis, 'structuredClone');
	const data = await new PptxHandler().load(bytes);
	const clones = clone.mock.calls.length;
	clone.mockRestore();
	const charts = data.slides[0]!.elements.filter(
		(element): element is ChartPptxElement => element.type === 'chart',
	);
	return { clones, charts };
}

describe('chart load work', () => {
	it('reads the package relationships once for all charts, not once per chart', async () => {
		const one = await load(await deck(1, 60));
		const eight = await load(await deck(8, 60));
		expect(eight.charts).toHaveLength(8);
		expect(eight.charts.every((chart) => chart.chartData?.dataLabelTextEditable)).toBe(true);
		// Seven more charts add their own reads, never another pass over the 60 parts.
		expect(eight.clones - one.clones).toBeLessThan(60);
	});

	it('leaves the embedded workbook unopened when the chart caches its data', async () => {
		const { charts } = await load(await deck(1, 0));
		const data = charts[0]!.chartData!;
		expect(data.categories).toEqual(['Q1', 'Q2']);
		expect(data.embeddedWorkbookData).toBeUndefined();
		expect(data.externalData?.embeddedWorkbookData?.byteLength).toBeGreaterThan(0);
	});
});
