import { readFileSync } from 'node:fs';
import { PptxHandler, type ChartPptxElement } from 'ooxml-core/pptx';
import { expect, it } from 'vitest';
import { buildChartViewModel } from './chart-view-model';

// PowerPoint 16 shape.Export capture of this exact gallery: slide 11 has no inline value labels;
// slide 14 shows mean X marks with no authored cx:layoutPr. This is a bounded defaults check.
it('matches captured native defaults for the committed funnel and box source parts', async () => {
	const bytes = readFileSync(
		new URL('../../../../core/pptx/__tests__/fixtures/e2e/chart-gallery.pptx', import.meta.url),
	);
	const deck = await new PptxHandler().load(
		bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer,
	);
	if (!deck) throw new Error('The committed chart gallery did not load.');
	const charts = deck.slides
		.flatMap((s) => s.elements)
		.filter((e) => e.type === 'chart') as ChartPptxElement[];
	const funnel = charts.find((e) => e.chartData.chartType === 'funnel')!;
	const box = charts.find((e) => e.chartData.chartType === 'boxWhisker')!;
	expect(buildChartViewModel(funnel).dataLabels).toEqual([]);
	expect(buildChartViewModel(box).primitives.filter((p) => p.kind === 'path')).toHaveLength(6);
	const explicit = {
		...box,
		chartData: {
			...box.chartData,
			series: box.chartData.series.map((s) => ({
				...s,
				boxWhiskerOptions: { showMeanMarker: false },
			})),
		},
	};
	expect(buildChartViewModel(explicit).primitives.filter((p) => p.kind === 'path')).toHaveLength(0);
});
