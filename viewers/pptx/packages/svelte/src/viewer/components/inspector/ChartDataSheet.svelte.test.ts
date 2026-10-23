import type { PptxChartData } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';

import ChartDataSheet from './ChartDataSheet.svelte';

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function chartData(): PptxChartData {
	return {
		chartType: 'bar',
		categories: ['Q1', 'Q2'],
		series: [
			{ name: 'Revenue', values: [10, 20] },
			{ name: 'Cost', values: [5, 6] },
		],
	};
}

function mountGrid(
	canEdit = true,
	data: PptxChartData = chartData(),
): { root: ShadowRoot; onreplace: ReturnType<typeof vi.fn> } {
	const onreplace = vi.fn();
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ChartDataSheet, { target, props: { data, canEdit, onreplace } });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return { root: target.querySelector('pptx-ui-chart-data-sheet')!.shadowRoot!, onreplace };
}

function byLabel(root: ShadowRoot, label: string): HTMLInputElement {
	return root.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
}

function setValue(control: HTMLInputElement, value: string): void {
	control.value = value;
	control.dispatchEvent(new Event('change', { bubbles: true }));
	flushSync();
}

describe('chartDataSheet (svelte)', () => {
	it('renders a value cell per category x series with the shared labels', () => {
		const { root } = mountGrid();

		expect(root.querySelectorAll('tbody tr')).toHaveLength(2);
		expect(byLabel(root, 'Revenue value 1')).not.toBeNull();
		expect(byLabel(root, 'Cost value 2')).not.toBeNull();
	});

	it('writes a single cell without disturbing the rest of the row', () => {
		const { root, onreplace } = mountGrid();

		setValue(byLabel(root, 'Revenue value 2'), '42');

		const next = onreplace.mock.calls[0][0] as PptxChartData;
		expect(next.series[0].values).toStrictEqual([10, 42]);
		expect(next.series[1].values).toStrictEqual([5, 6]);
	});

	it('ignores a cleared cell rather than writing zero', () => {
		const { root, onreplace } = mountGrid();

		setValue(byLabel(root, 'Revenue value 1'), '');

		expect(onreplace).not.toHaveBeenCalled();
	});

	it('adds a category and a series through the header buttons', () => {
		const { root, onreplace } = mountGrid();

		root.querySelector<HTMLButtonElement>('[aria-label="Add category"]')?.click();
		flushSync();
		expect((onreplace.mock.calls[0][0] as PptxChartData).categories).toStrictEqual([
			'Q1',
			'Q2',
			'Cat 3',
		]);

		root.querySelector<HTMLButtonElement>('[aria-label="Add series"]')?.click();
		flushSync();
		expect((onreplace.mock.calls[1][0] as PptxChartData).series).toHaveLength(3);
	});

	it('renames a series and keeps its styling', () => {
		const styled = {
			...chartData(),
			series: [{ name: 'Revenue', values: [1, 2], color: '#123456' }],
		};
		const { root, onreplace } = mountGrid(true, styled);

		setValue(byLabel(root, 'Series 1'), 'Net revenue');

		expect((onreplace.mock.calls[0][0] as PptxChartData).series[0]).toMatchObject({
			name: 'Net revenue',
			color: '#123456',
		});
	});

	it('hides the remove button when only one series or category is left', () => {
		const single: PptxChartData = {
			chartType: 'bar',
			categories: ['Only'],
			series: [{ name: 'Solo', values: [1] }],
		};
		const { root } = mountGrid(true, single);

		expect(root.querySelector('button.remove')).toBeNull();
	});

	it('hides add/remove and disables every input in a read-only viewer', () => {
		const { root } = mountGrid(false);

		expect(root.querySelector('.actions')).toBeNull();
		expect(root.querySelector('button.remove')).toBeNull();
		expect(
			Array.from(root.querySelectorAll<HTMLInputElement>('input')).every((input) => input.disabled),
		).toBeTruthy();
	});
});
