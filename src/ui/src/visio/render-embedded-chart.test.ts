import { afterEach, describe, expect, it } from 'vitest';
import type { ChartExFamily } from 'ooxml-core/chart';
import type { VisioShape } from 'ooxml-core/visio';
import type { CellValue } from 'ooxml-core/xlsx';
import { gridChartWorkbook, parseChartGrid, readEmbeddedChart } from 'ooxml-core/xlsx/ui';
import { renderEmbeddedChart } from './render-embedded-chart';
import { gridChartExWorkbook } from './viewer-insert-chart-ex';

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

	it('is an image a reader can name and whose data it can read', async () => {
		const plain = renderEmbeddedChart(await chartShape())!;
		const svg = plain.querySelector('svg')!;
		expect(svg.getAttribute('role')).toBe('img');
		// The chart's own title names it, and its series and values are its description.
		expect(svg.querySelector('title')?.textContent).toBe('Quarter one');
		expect(svg.querySelector('desc')?.textContent).toContain('Jan 10; Feb 20');
		// The shape's alt text is what the author wrote for a reader, so it names the chart.
		const described = await chartShape();
		described.alternativeText = 'Sales doubled from January to February';
		expect(renderEmbeddedChart(described)!.querySelector('svg > title')?.textContent).toBe(
			'Sales doubled from January to February',
		);
	});

	it('keeps script out of anything the chart carries', async () => {
		const shape = await chartShape();
		shape.embeddedObject!.chart.title = '<script>alert(1)</script>';
		const group = renderEmbeddedChart(shape)!;
		expect(group.querySelector('script')).toBeNull();
		expect(group.innerHTML).not.toContain('<script');
	});
});

const CHART_EX_CASES: ReadonlyArray<
	readonly [family: ChartExFamily, expectedType: string, rows: readonly (readonly CellValue[])[]]
> = [
	[
		'waterfall',
		'waterfall',
		[
			['Item', 'Amount'],
			['Start', 100],
			['Sales', 60],
			['Costs', -40],
		],
	],
	[
		'clusteredColumn',
		'histogram',
		[
			['Sample', 'Score'],
			['A', 10],
			['B', 14],
			['C', 18],
			['D', 22],
			['E', 28],
		],
	],
	[
		'pareto',
		'histogram',
		[
			['Defect', 'Count'],
			['Scratch', 40],
			['Dent', 25],
			['Crack', 10],
		],
	],
	[
		'boxWhisker',
		'boxWhisker',
		[
			['Group', 'S1'],
			['A', 10],
			['A', 15],
			['A', 22],
			['B', 8],
			['B', 18],
			['B', 26],
		],
	],
	[
		'treemap',
		'treemap',
		[
			['Branch', 'Leaf', 'Value'],
			['Fruit', 'Apple', 30],
			['Fruit', 'Pear', 20],
			['Veg', 'Kale', 25],
		],
	],
	[
		'sunburst',
		'sunburst',
		[
			['Branch', 'Leaf', 'Value'],
			['Fruit', 'Apple', 30],
			['Fruit', 'Pear', 20],
			['Veg', 'Kale', 25],
		],
	],
	[
		'funnel',
		'funnel',
		[
			['Stage', 'Count'],
			['Visits', 1000],
			['Trials', 240],
			['Paid', 80],
		],
	],
	[
		'regionMap',
		'regionMap',
		[
			['Country', 'Sales'],
			['United States', 120],
			['Canada', 80],
			['France', 60],
		],
	],
];

describe('embedded ChartEx drawing', () => {
	it.each(CHART_EX_CASES)(
		'renders embedded %s ChartEx via the shared SVG painter',
		async (family, expectedType, rows) => {
			const title = `ChartEx ${family}`;
			const bytes = (await gridChartExWorkbook(rows, family, { title }))!;
			const embedded = (await readEmbeddedChart(bytes))!;
			const shape = {
				id: '10',
				kind: 'foreign',
				width: 6.4,
				height: 4,
				embeddedObject: { bytes, chart: embedded.view, chartEx: true },
			} as unknown as VisioShape;
			const group = renderEmbeddedChart(shape)!;
			expect(group.hasAttribute('data-embedded-chart')).toBe(true);
			expect(group.hasAttribute('data-embedded-chart-ex')).toBe(true);
			expect(group.dataset.chartType).toBe(expectedType);
			const svg = group.querySelector('svg')!;
			expect(svg.getAttribute('role')).toBe('img');
			expect(svg.querySelector('title')?.textContent).toBe(title);
			expect(group.textContent).toContain(title);
			expect(group.textContent).not.toContain('not drawn');
			if (family === 'sunburst') {
				expect(group.querySelectorAll('path').length).toBeGreaterThan(0);
			} else if (family === 'regionMap') {
				expect(group.querySelector('[data-chart-part="regionMap"]')).not.toBeNull();
				expect(group.querySelectorAll('path').length).toBeGreaterThan(0);
				expect(group.textContent).toContain('approximate');
			} else {
				expect(group.querySelectorAll('rect').length).toBeGreaterThan(0);
			}
			if (family === 'pareto') {
				expect(group.querySelectorAll('polyline').length).toBeGreaterThan(0);
			}
		},
	);
});
