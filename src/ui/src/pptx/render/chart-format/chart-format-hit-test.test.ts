// @vitest-environment jsdom
import type { ChartPptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { computeChartLegendLayout } from '../chart-legend-layout';
import { buildChartViewModel } from '../chart-view-model';
import { chartFormatTargetAtPoint, chartFormatTargetFromEvent } from './chart-format-hit-test';
import { columnChart } from './chart-format-fixtures.test-util';

function chartElement(): ChartPptxElement {
	const chartData = columnChart();
	chartData.axes = chartData.axes!.slice(0, 2);
	return {
		id: 'c1',
		type: 'chart',
		x: 0,
		y: 0,
		width: 480,
		height: 300,
		chartData,
	} as ChartPptxElement;
}

describe('chartFormatTargetAtPoint', () => {
	const element = chartElement();
	const vm = buildChartViewModel(element);
	const data = element.chartData!;

	it('finds the legend, axis labels, gridlines, plot area and chart area', () => {
		const legend = computeChartLegendLayout(vm)[0]!;
		expect(chartFormatTargetAtPoint(vm, data, { x: legend.x + 4, y: legend.y - 2 })).toBe('legend');
		const category = vm.categoryLabels[0]!;
		expect(chartFormatTargetAtPoint(vm, data, { x: category.x, y: category.y - 3 })).toBe('axis:0');
		const value = vm.axisLabels[0]!;
		expect(
			chartFormatTargetAtPoint(vm, data, {
				x: value.textAnchor === 'end' ? value.x - 3 : value.x + 3,
				y: value.y - 3,
			}),
		).toBe('axis:1');
		const line = vm.gridlines[1]!;
		expect(chartFormatTargetAtPoint(vm, data, { x: (line.x1 + line.x2) / 2, y: line.y1 + 1 })).toBe(
			'majorGridlines:1',
		);
		const top = Math.min(...vm.gridlines.map((l) => l.y1));
		const bottom = Math.max(...vm.gridlines.map((l) => l.y1));
		const between = vm.gridlines
			.map((l) => l.y1)
			.sort((a, b) => a - b)
			.slice(0, 2);
		expect(
			chartFormatTargetAtPoint(vm, data, {
				x: (line.x1 + line.x2) / 2,
				y: (between[0]! + between[1]!) / 2,
			}),
		).toBe('plotArea');
		expect(top).toBeLessThan(bottom);
		expect(chartFormatTargetAtPoint(vm, data, { x: 1, y: 1 })).toBe('chartArea');
	});
});

describe('chartFormatTargetFromEvent', () => {
	it('reads series and title tags off the event path before geometry', () => {
		const host = document.createElement('div');
		host.innerHTML =
			'<svg viewBox="0 0 480 300"><rect data-chart-part="dataPoint" data-chart-series="1" data-chart-point="0"></rect><text data-chart-part="title">T</text></svg>';
		document.body.append(host);
		try {
			const mark = host.querySelector('rect')!;
			const title = host.querySelector('text')!;
			const at = (target: Element) => ({ target, clientX: 0, clientY: 0 });
			expect(chartFormatTargetFromEvent(at(mark), host, chartElement())).toBe('series:1');
			expect(chartFormatTargetFromEvent(at(title), host, chartElement())).toBe('title');
			expect(chartFormatTargetFromEvent(at(host), host, null)).toBe('chartArea');
		} finally {
			host.remove();
		}
	});
});
