// @vitest-environment happy-dom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import type { PptxChartData } from 'ooxml-core/pptx';

import { registerPptxWebControls } from './index';
import type { ChartDataSheetChange, PptxUiChartDataSheetElement } from './chart-data-sheet';

beforeAll(() => registerPptxWebControls());
afterEach(() => document.body.replaceChildren());

const chartData = {
	chartType: 'bar',
	categories: ['Q1', 'Q2'],
	series: [
		{ name: 'North', values: [1, 2], color: '#FF0000' },
		{ name: 'South', values: [3] },
	],
} as PptxChartData;

function mount(state: Partial<PptxUiChartDataSheetElement['state']> = {}) {
	const sheet = document.createElement('pptx-ui-chart-data-sheet');
	document.body.append(sheet);
	sheet.state = { chartData, canEdit: true, ...state };
	const seen: ChartDataSheetChange[] = [];
	sheet.addEventListener('pptx-chart-data-change', (event) =>
		seen.push((event as CustomEvent<ChartDataSheetChange>).detail),
	);
	return { sheet, seen };
}
const cell = (sheet: HTMLElement, row: number, col: number) =>
	sheet.shadowRoot!.querySelector<HTMLInputElement>(`input[data-row="${row}"][data-col="${col}"]`)!;
const type = (input: HTMLInputElement, value: string) => {
	input.value = value;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('pptx-ui-chart-data-sheet', () => {
	it('shows the chart data with the pptx labels', () => {
		const t = (key: string, params?: Record<string, string | number>) =>
			params ? `${key}:${params.name}:${params.number}` : `T(${key})`;
		const { sheet } = mount({ translate: t });
		expect(sheet.shadowRoot!.querySelector('.heading')!.textContent).toBe('T(pptx.chart.data)');
		expect(cell(sheet, 1, 1).getAttribute('aria-label')).toBe('pptx.chart.seriesValueCell:North:1');
		// A ragged series reads as a blank cell.
		expect(cell(sheet, 2, 2).value).toBe('');
	});

	it('commits edits through the grid policy, keeping series formatting', () => {
		const { sheet, seen } = mount();
		type(cell(sheet, 2, 2), '8');
		expect(seen[0]!.chartData.series[1]!.values).toEqual([3, 8]);
		type(cell(sheet, 0, 1), 'East');
		expect(seen[1]!.chartData.series[0]).toMatchObject({ name: 'East', color: '#FF0000' });
		expect(seen[1]!.chartData.series[1]!.values).toEqual([3, 8]);
	});

	it('honours the follow-data-point option and marks the canvas selection', () => {
		const withPoints = {
			...chartData,
			series: [{ name: 'North', values: [1, 2], dataPoints: [{ idx: 1, color: '#00FF00' }] }],
		} as unknown as PptxChartData;
		const { sheet, seen } = mount({
			chartData: withPoints,
			followDataPoint: false,
			highlight: { seriesIndex: 0, pointIndex: 1 },
		});
		expect(cell(sheet, 2, 1).classList.contains('highlight')).toBe(true);
		sheet.shadowRoot!.querySelectorAll<HTMLButtonElement>('button.remove')[0]!.click();
		expect(seen[0]!.chartData.categories).toEqual(['Q2']);
	});

	it('is hidden without chart data and read-only without edit rights', () => {
		const { sheet, seen } = mount({ canEdit: false });
		expect(cell(sheet, 1, 1).disabled).toBe(true);
		type(cell(sheet, 1, 1), '5');
		expect(seen).toHaveLength(0);
		sheet.state = { chartData: undefined, canEdit: true };
		expect(sheet.hidden).toBe(true);
	});
});
