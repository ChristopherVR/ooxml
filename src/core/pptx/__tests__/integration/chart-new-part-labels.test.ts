/**
 * Data label positions of a NEW chart: the SDK refuses a position the chart type does not offer
 * (PowerPoint repairs a part with one), an offered position is written and reloads, and a refused
 * position set on the model directly is left out of the saved part with a save warning.
 */
import { describe, expect, it } from 'vitest';

import {
	setChartDataLabels,
	setChartDataPointLabel,
	setChartSeriesChartType,
} from '../../core/builders/sdk/chart-operations';
import { withChartDataLabelOptions } from '../../core/builders/sdk/chart-format-options-operations';
import { firstElement, newChartDeck } from '../helpers/new-chart-deck';

describe('new chart label positions: SDK checks', () => {
	it('refuses a chart-wide position the type does not offer', async () => {
		const { element } = await newChartDeck('line');
		expect(() => setChartDataLabels(element, { showValue: true, position: 'outEnd' })).toThrow(
			RangeError,
		);
		expect(element.chartData!.style?.dataLabels).toBeUndefined();
		const area = await newChartDeck('area');
		expect(() => setChartDataLabels(area.element, { position: 'ctr' })).toThrow(/takes none/u);
		const doughnut = await newChartDeck('doughnut');
		expect(() =>
			withChartDataLabelOptions(doughnut.element.chartData!, { position: 'bestFit' }),
		).toThrow(RangeError);
	});

	it("refuses a series or point position its group does not offer, stacked bars' outEnd too", async () => {
		const { element } = await newChartDeck('bar', { grouping: 'stacked' });
		expect(() =>
			withChartDataLabelOptions(element.chartData!, { showValue: true, position: 'outEnd' }, 1),
		).toThrow(RangeError);
		expect(() => setChartDataPointLabel(element, 0, 1, { position: 'outEnd' })).toThrow(RangeError);
		expect(() => setChartDataPointLabel(element, 0, 1, { position: 'inBase' })).not.toThrow();
	});

	it('checks a combination chart against the group each series is in', async () => {
		const { element } = await newChartDeck('bar');
		setChartSeriesChartType(element, 1, 'line');
		expect(() => setChartDataLabels(element, { position: 'inEnd' })).toThrow(RangeError);
		expect(() =>
			withChartDataLabelOptions(element.chartData!, { showValue: true, position: 't' }, 1),
		).not.toThrow();
		expect(() =>
			withChartDataLabelOptions(element.chartData!, { showValue: true, position: 't' }, 0),
		).toThrow(RangeError);
	});

	it('accepts clearing a position on any type', async () => {
		const { element } = await newChartDeck('area');
		expect(() =>
			setChartDataLabels(element, { showValue: true, position: undefined }),
		).not.toThrow();
		expect(() => withChartDataLabelOptions(element.chartData!, { position: null })).not.toThrow();
	});
});

describe('new chart label positions: save and reload', () => {
	it('writes and reloads an offered series position', async () => {
		const { element, save } = await newChartDeck('bar');
		element.chartData = withChartDataLabelOptions(
			element.chartData!,
			{ showValue: true, position: 'outEnd' },
			1,
		);
		const { xml, chart, warnings } = await save();
		const second = firstElement(xml, 'c:ser', xml.indexOf('<c:ser>') + 1);
		expect(second).toContain('<c:dLblPos val="outEnd"></c:dLblPos>');
		expect(firstElement(xml, 'c:ser')).not.toContain('dLblPos');
		expect(chart.series[1]!.dataLabelOptions).toMatchObject({
			showValue: true,
			position: 'outEnd',
		});
		expect(chart.series[0]!.dataLabelOptions?.showValue).not.toBe(true);
		expect(warnings).toStrictEqual([]);
	});

	it('writes labels with no position when none is set', async () => {
		const { element, save } = await newChartDeck('line');
		setChartDataLabels(element, { showValue: true });
		const { xml, chart } = await save();
		expect(xml).toContain('<c:showVal val="1"></c:showVal>');
		expect(xml).not.toContain('dLblPos');
		expect(chart.style?.hasDataLabels).toBe(true);
		expect(chart.style?.dataLabels?.position).toBeUndefined();
	});

	it('writes a pie best fit and a line above', async () => {
		const pie = await newChartDeck('pie');
		setChartDataLabels(pie.element, { showPercent: true, position: 'bestFit' });
		expect((await pie.save()).chart.style?.dataLabels?.position).toBe('bestFit');
		const line = await newChartDeck('line');
		setChartDataPointLabel(line.element, 2, 0, { showValue: true, position: 't' });
		const saved = await line.save();
		expect(saved.chart.series[2]!.dataLabels?.find((label) => label.idx === 0)?.position).toBe('t');
	});

	it('leaves out, and reports, a refused position set on the model directly', async () => {
		const { element, save } = await newChartDeck('line');
		element.chartData!.style = {
			...element.chartData!.style,
			hasDataLabels: true,
			dataLabels: { showValue: true, position: 'outEnd' },
		};
		element.chartData!.series[0]!.dataLabelOptions = { showValue: true, position: 'inBase' };
		const { xml, chart, warnings } = await save();
		expect(xml).not.toContain('dLblPos');
		expect(chart.style?.dataLabels?.showValue).toBe(true);
		expect(chart.series[0]!.dataLabelOptions?.showValue).toBe(true);
		expect(warnings.map((warning) => warning.code)).toStrictEqual(['CHART_FIELDS_NOT_WRITTEN']);
		expect(warnings[0]!.message).toContain('style.dataLabels.position');
		expect(warnings[0]!.message).toContain('series.0.dataLabelOptions.position');
		// The model the caller holds is not changed by the save.
		expect(element.chartData!.style?.dataLabels?.position).toBe('outEnd');
	});
});
