import { afterEach, describe, expect, it } from 'vitest';
import type { VisioShape } from 'ooxml-core/visio';
import { gridChartWorkbook, parseChartGrid, readEmbeddedChart } from 'ooxml-core/xlsx/ui';
import { renderEmbeddedChart } from './render-embedded-chart';

afterEach(() => document.body.replaceChildren());

async function chartShape(patch: Partial<VisioShape['embeddedObject'] & object> = {}) {
	const bytes = (await gridChartWorkbook(
		parseChartGrid('Month\tSales\nJan\t10\nFeb\t20'),
		'column',
		{
			title: 'Quarter one',
		},
	))!;
	const embedded = (await readEmbeddedChart(bytes))!;
	return {
		id: '1',
		kind: 'foreign',
		width: 6.4,
		height: 4,
		embeddedObject: { bytes, chart: embedded.view, ...patch },
	} as unknown as VisioShape;
}

describe('embedded chart drawing', () => {
	it('draws the chart from its workbook into the shape frame, flipped to read upright', async () => {
		const group = renderEmbeddedChart(await chartShape({ x: 0, y: 0, width: 6.4, height: 4 }))!;
		expect(group.hasAttribute('data-embedded-chart')).toBe(true);
		const svg = group.querySelector('svg')!;
		expect(svg.getAttribute('width')).toBe('6.4');
		expect(svg.getAttribute('height')).toBe('4');
		expect(svg.getAttribute('transform')).toBe('translate(0 4) scale(1 -1)');
		expect(svg.getAttribute('viewBox')).toMatch(/^0 0 \d+ \d+$/);
		expect(group.textContent).toContain('Quarter one');
		expect(group.querySelectorAll('rect').length).toBeGreaterThan(0);
	});

	it('scales into a resized frame and honours opacity and placement', async () => {
		const group = renderEmbeddedChart(
			await chartShape({ x: 1, y: 2, width: 3, height: 2, opacity: 0.5 }),
		)!;
		const svg = group.querySelector('svg')!;
		expect(svg.getAttribute('transform')).toBe('translate(1 4) scale(1 -1)');
		expect(svg.getAttribute('opacity')).toBe('0.5');
		expect(svg.getAttribute('preserveAspectRatio')).toBe('none');
	});

	it('draws nothing for a shape without a chart or with no size', async () => {
		expect(renderEmbeddedChart({ id: '2', width: 1, height: 1 } as VisioShape)).toBeUndefined();
		expect(renderEmbeddedChart(await chartShape({ width: 0, height: 0 }))).toBeUndefined();
	});

	it('keeps script out of anything the chart carries', async () => {
		const shape = await chartShape();
		shape.embeddedObject!.chart.title = '<script>alert(1)</script>';
		const group = renderEmbeddedChart(shape)!;
		expect(group.querySelector('script')).toBeNull();
		expect(group.innerHTML).not.toContain('<script');
	});
});
