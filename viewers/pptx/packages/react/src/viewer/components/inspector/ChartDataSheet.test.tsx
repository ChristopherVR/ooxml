// @vitest-environment happy-dom
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import type { PptxChartData } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { ChartDataSheet } from './ChartDataSheet';

vi.mock<typeof import('../../i18n/use-translation')>(import('../../i18n/use-translation'), () => ({
	useTranslation: () => ({ t: (key: string) => translationsEn[key] ?? key }),
}));

beforeAll(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	registerPptxWebControls();
});

const mounted: { root: Root; target: HTMLElement }[] = [];
afterEach(() => {
	for (const { root, target } of mounted.splice(0)) {
		act(() => root.unmount());
		target.remove();
	}
});

const chartData = {
	chartType: 'bar',
	categories: ['Q1', 'Q2'],
	series: [{ name: 'North', values: [1, 2], color: '#FF0000' }],
} as PptxChartData;

function render(props: Partial<React.ComponentProps<typeof ChartDataSheet>> = {}) {
	const onChange = vi.fn();
	const target = document.createElement('div');
	document.body.append(target);
	const root = createRoot(target);
	mounted.push({ root, target });
	act(() =>
		root.render(
			<ChartDataSheet
				chartData={chartData}
				canEdit
				followDataPoint
				onChange={onChange}
				{...props}
			/>,
		),
	);
	const sheet = target.querySelector('pptx-ui-chart-data-sheet')!;
	const cell = (row: number, col: number) =>
		sheet.shadowRoot!.querySelector<HTMLInputElement>(
			`input[data-row="${row}"][data-col="${col}"]`,
		)!;
	return { onChange, sheet, cell };
}

describe('chartDataSheet (react)', () => {
	it('feeds the shared sheet the chart and commits its edits', () => {
		const { onChange, cell } = render();
		expect(cell(2, 1).value).toBe('2');
		expect(cell(1, 1).getAttribute('aria-label')).toBe('North value 1');
		cell(2, 1).value = '9';
		cell(2, 1).dispatchEvent(new Event('change', { bubbles: true }));
		expect(onChange).toHaveBeenCalledOnce();
		const next = onChange.mock.calls[0]![0] as PptxChartData;
		expect(next.series[0]).toMatchObject({ values: [1, 9], color: '#FF0000' });
	});

	it('marks the canvas-selected cell and is read-only without edit rights', () => {
		const { cell, sheet } = render({
			canEdit: false,
			highlightCell: { seriesIndex: 0, pointIndex: 1 },
		});
		expect(cell(2, 1).classList.contains('highlight')).toBeTruthy();
		expect(cell(2, 1).disabled).toBeTruthy();
		expect(sheet.shadowRoot!.querySelector('[data-action="add-series"]')).toBeNull();
	});
});
