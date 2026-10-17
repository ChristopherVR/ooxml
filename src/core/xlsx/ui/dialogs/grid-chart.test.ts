import { describe, expect, it } from 'vitest';
import {
	GROUPED_CHART_TYPES,
	INSERT_CHART_TYPES,
	MAX_GRID_COLS,
	MAX_GRID_ROWS,
	gridChart,
	gridChartSvg,
	parseChartGrid,
} from './grid-chart';

describe('parseChartGrid', () => {
	it('reads tab or comma separated rows, numbers as numbers and short rows padded', () => {
		expect(parseChartGrid('Month\tSales\nJan\t10\nFeb')).toEqual([
			['Month', 'Sales'],
			['Jan', 10],
			['Feb', null],
		]);
		expect(parseChartGrid('a, b\r\n1, 2.5\n\n')).toEqual([
			['a', 'b'],
			[1, 2.5],
		]);
	});

	it('caps the grid so pasted text cannot build an enormous chart', () => {
		const wide = Array.from({ length: MAX_GRID_COLS + 10 }, (_, i) => i).join(',');
		const tall = Array.from({ length: MAX_GRID_ROWS + 10 }, () => wide).join('\n');
		const grid = parseChartGrid(tall);
		expect(grid).toHaveLength(MAX_GRID_ROWS);
		expect(grid[0]).toHaveLength(MAX_GRID_COLS);
	});
});

describe('gridChart', () => {
	const rows = parseChartGrid('Month\tSales\tCost\nJan\t10\t4\nFeb\t20\t8\nMar\t30\t9');

	it('names series from the header row and labels categories from the first column', () => {
		const grid = gridChart(rows, 'column', { title: 'Q1' })!;
		expect(grid.chart.title).toBe('Q1');
		expect(grid.chart.series.map((s) => s.name)).toEqual(['Sales', 'Cost']);
		expect(grid.chart.series[0]!.categories).toEqual(['Jan', 'Feb', 'Mar']);
	});

	it('has no chart for an empty grid', () => {
		expect(gridChart([], 'column')).toBeUndefined();
		expect(gridChart([[]], 'column')).toBeUndefined();
	});

	it('draws every insertable type as an SVG document', () => {
		for (const [type] of INSERT_CHART_TYPES) {
			const svg = gridChartSvg(gridChart(rows, type)!, 360, 220);
			expect(svg.startsWith('<svg'), type).toBe(true);
		}
		const bare = gridChartSvg(gridChart(rows, 'column', { title: 'Hidden' })!, 96, 60, true);
		expect(bare).not.toContain('Hidden');
	});

	it('knows which types can be grouped', () => {
		expect(GROUPED_CHART_TYPES).toEqual(['column', 'bar']);
	});
});
