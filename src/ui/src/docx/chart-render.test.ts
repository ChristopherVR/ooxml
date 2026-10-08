// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { chartsIn, loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer, layoutDocumentModel } from 'ooxml-core/docx/layout';
import type { EditorView } from 'prosemirror-view';
import { ImageMediaCache, imageNodeView } from './image-media';
import { renderPrintLayout } from './print-layout';
import { runToInlineNodes } from './run-adapter';

// Built by ooxml-core (scripts/docx/build-chart-fixture.mjs) from the Excel chart of
// excel-features.xlsx: one inline clustered column chart, two series of four regions.
const fixture = () =>
	new Uint8Array(readFileSync(join(process.cwd(), 'src/docx/__fixtures__/chart.docx')));

async function chartRun() {
	const { model } = await loadDocx(fixture());
	const run = model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.filter((r) => r.image?.chart) : [],
	)[0]!;
	return { model, run };
}

describe('Word charts in the editor', () => {
	it('draws the chart as SVG in the image node view, with the document theme', async () => {
		const { model, run } = await chartRun();
		const node = runToInlineNodes(run)[0]!;
		const view = imageNodeView(new ImageMediaCache(() => undefined), { theme: () => model.theme })(
			node,
			{ editable: true } as unknown as EditorView,
			() => 0,
		);
		const dom = view.dom as HTMLElement;
		expect(dom.dataset.docxChart).toBe('1');
		expect(dom.style.width).toBe(`${run.image!.widthPx}px`);
		const svg = dom.querySelector('svg')!;
		expect(svg.namespaceURI).toBe('http://www.w3.org/2000/svg');
		expect(svg.querySelectorAll('g[data-chart-series="0"]')).toHaveLength(4);
		expect(svg.querySelectorAll('g[data-chart-series="1"]')).toHaveLength(4);
		const fill = svg.querySelector('g[data-chart-series="0"] > *')!.getAttribute('fill');
		expect(fill).toBe(`#${model.theme!.colors.accent1!.toUpperCase()}`);
		// Read-only: the node view ignores DOM mutations and the chart part is untouched.
		expect(view.ignoreMutation?.({} as never)).toBe(true);
		expect(chartsIn(model.blocks)).toHaveLength(1);
	});
});

describe('Word charts in Print Layout', () => {
	it('draws the chart as SVG on the page at the drawing extent', async () => {
		const { model, run } = await chartRun();
		const handle = renderPrintLayout(layoutDocumentModel(model, createFakeMeasurer()));
		const chart = handle.element.querySelector<HTMLElement>('[data-docx-chart="1"]')!;
		expect(chart).not.toBeNull();
		expect(chart.style.width).toBe(`${run.image!.widthPx}px`);
		expect(chart.style.height).toBe(`${run.image!.heightPx}px`);
		const svg = chart.querySelector('svg')!;
		expect(svg.namespaceURI).toBe('http://www.w3.org/2000/svg');
		expect(svg.querySelectorAll('g[data-chart-series="0"]')).toHaveLength(4);
		expect(svg.querySelectorAll('g[data-chart-series="1"]')).toHaveLength(4);
		expect(handle.element.querySelector('.dve-print-picture-missing')).toBeNull();
	});
});
