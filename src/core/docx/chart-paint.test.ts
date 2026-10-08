// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { DOMSerializer, Schema } from 'prosemirror-model';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx, type DocxChart } from './index';
import { docxChartPaint } from './chart-paint';
import { adaptDocumentModel } from './layout/adapter';
import { imageNodeSpec } from './ui/inline-content-schema';
import { runToInlineNodes } from './ui/run-adapter';
import { markSpecs } from './ui/schema-marks';

// Built by scripts/docx/build-chart-fixture.mjs from the Excel chart of excel-features.xlsx.
const fixture = () =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__/chart.docx')));

async function loadChart() {
	const loaded = await loadDocx(fixture());
	const chart = chartsIn(loaded.model.blocks)[0]!;
	return { loaded, chart };
}

const seriesMarks = (svg: string, series: number) =>
	svg.match(new RegExp(`<g data-chart-series="${series}"[^>]*><[a-z]+ [^>]*>`, 'g')) ?? [];

describe('Word charts are drawn', () => {
	it('reads the formatting and colour style the painter needs', async () => {
		const { chart } = await loadChart();
		expect(chart.formatting).toBeDefined();
		expect(chart.colorPalette).toBe(10);
		expect(chart.notice).toContain('not editable');
	});

	it('paints the chart with the document theme at the drawing extent', async () => {
		const { loaded, chart } = await loadChart();
		const paint = docxChartPaint(chart, 576, 336, loaded.model.theme);
		expect(paint.label).toBe('Chart: Sales by region');
		expect(paint.svg).toMatch(/^<svg [^>]*width="576" height="336"/);
		expect(seriesMarks(paint.svg!, 0)).toHaveLength(4);
		expect(seriesMarks(paint.svg!, 1)).toHaveLength(4);
		const accent1 = loaded.model.theme?.colors.accent1;
		expect(accent1).toBeTruthy();
		expect(seriesMarks(paint.svg!, 0)[0]).toContain(`fill="#${accent1!.toUpperCase()}"`);
	});

	it('shows a labelled frame for a family the painter does not draw', async () => {
		const { chart } = await loadChart();
		const bubble = structuredClone(chart) as DocxChart;
		bubble.chartSpace!.plotArea.groups[0]!.kind = 'bubble';
		expect(docxChartPaint(bubble, 100, 100)).toEqual({ label: 'Chart (bubble)' });
		const { chartSpace: _space, ...unread } = chart;
		expect(docxChartPaint(unread, 100, 100)).toEqual({ label: 'Chart' });
	});

	it('renders an SVG in the editor DOM spec, sized to the extent', async () => {
		const { loaded } = await loadChart();
		const run = loaded.model.blocks.flatMap((block) =>
			block.type === 'paragraph' ? block.runs.filter((r) => r.image?.chart) : [],
		)[0]!;
		const schema = new Schema({
			nodes: { doc: { content: 'inline*' }, text: { group: 'inline' }, image: imageNodeSpec },
			marks: markSpecs,
		});
		const [node] = runToInlineNodes(run, schema);
		const { dom } = DOMSerializer.renderSpec(document, imageNodeSpec.toDOM!(node!));
		const span = dom as HTMLElement;
		expect(span.dataset.docxChart).toBe('1');
		expect(span.style.width).toBe('576px');
		expect(span.style.height).toBe('336px');
		const svg = span.querySelector('svg')!;
		expect(svg.namespaceURI).toBe('http://www.w3.org/2000/svg');
		expect(svg.querySelectorAll('g[data-chart-series="0"]')).toHaveLength(4);
		expect(svg.querySelectorAll('g[data-chart-series="1"]')).toHaveLength(4);
	});

	it('carries the painted chart to the print layout input', async () => {
		const { loaded } = await loadChart();
		const input = adaptDocumentModel(loaded.model);
		const objects = JSON.stringify(input).match(/"chart":\{"svg":"<svg /g) ?? [];
		expect(objects).toHaveLength(1);
	});
});
