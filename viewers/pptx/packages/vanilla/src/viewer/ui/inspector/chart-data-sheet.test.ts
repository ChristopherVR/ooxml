import type { PptxChartData } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createChartDataSheet } from './chart-data-sheet';

afterEach(() => document.body.replaceChildren());

const chartData = (): PptxChartData => ({
	chartType: 'bar',
	categories: ['A', 'B'],
	series: [
		{ name: 'S1', values: [1, 2], color: '#112233' },
		{ name: 'S2', values: [3, 4] },
	],
});

function mount(getFollowDataPoint?: () => boolean) {
	const onChange = vi.fn();
	const sheet = createChartDataSheet(document, createTranslator(), onChange, getFollowDataPoint);
	document.body.append(sheet.el);
	sheet.update(chartData());
	const root = sheet.el.shadowRoot!;
	const cell = (row: number, col: number) =>
		root.querySelector<HTMLInputElement>(`input[data-row="${row}"][data-col="${col}"]`)!;
	return { sheet, onChange, root, cell };
}

const commit = (input: HTMLInputElement, value: string) => {
	input.value = value;
	input.dispatchEvent(new Event('change', { bubbles: true }));
};

describe('chart data sheet (vanilla)', () => {
	it('shows one column per series and one row per category', () => {
		const { root, cell } = mount();
		expect(root.querySelectorAll('thead th')).toHaveLength(3);
		expect(root.querySelectorAll('tbody tr')).toHaveLength(2);
		expect(cell(1, 1).getAttribute('aria-label')).toBe('S1 value 1');
	});

	it('commits an edited cell, keeping series formatting, and refuses non-numbers', () => {
		const { onChange, cell } = mount();
		commit(cell(1, 1), '');
		expect(onChange).not.toHaveBeenCalled();
		commit(cell(1, 1), '42');
		expect(onChange.mock.calls[0]![0].series[0]).toMatchObject({
			values: [42, 2],
			color: '#112233',
		});
	});

	it('removes a category with the follow-data-point option and never the last series', () => {
		const follow = vi.fn(() => false);
		const { sheet, onChange, root } = mount(follow);
		root.querySelector<HTMLButtonElement>('th.category button.remove')!.click();
		expect(onChange.mock.calls[0]![0].categories).toEqual(['B']);
		expect(follow).toHaveBeenCalled();
		sheet.update({ chartType: 'bar', categories: ['A'], series: [{ name: 'S1', values: [1] }] });
		expect(root.querySelector('button.remove')).toBeNull();
	});

	it('marks the canvas selection and goes read-only', () => {
		const { sheet, root, cell } = mount();
		sheet.update(chartData(), { seriesIndex: 1, pointIndex: 0 });
		expect(cell(1, 2).classList.contains('highlight')).toBe(true);
		sheet.setEditable(false);
		expect(cell(1, 1).disabled).toBe(true);
		expect(root.querySelector('[data-action="add-series"]')).toBeNull();
		sheet.update(undefined);
		expect(sheet.el.hidden).toBe(true);
	});
});
