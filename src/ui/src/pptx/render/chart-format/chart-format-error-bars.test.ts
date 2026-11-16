/**
 * PowerPoint's Error Bars section in the shared chart Format pane: every
 * binding mounts `pptx-ui-chart-format-pane`, which reads the chart through
 * `pptxChartFormatSource` and applies a field through `applyChartFormatField`,
 * so these cover the React, Vue, Svelte, Angular and vanilla editors alike.
 */
import { describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n';
import { applyChartFormatField, buildChartFormatPaneModel, pptxChartFormatKey } from '.';
import { columnChart, pieChart } from './chart-format-fixtures.test-util';

const seriesOptions = (data = columnChart()) =>
	buildChartFormatPaneModel(data, 'series:0', 'options');

describe('PowerPoint chart error bars', () => {
	it('shows the Trendlines and Error Bars sections for 2-D unstacked series only', () => {
		const sections = seriesOptions().sections.map((section) => section.id);
		expect(sections).toContain('errorBars');
		expect(sections).toContain('trendlines');
		expect(seriesOptions(pieChart()).sections.map((section) => section.id)).not.toContain(
			'errorBars',
		);
		const stacked = { ...columnChart(), grouping: 'stacked' as const };
		expect(seriesOptions(stacked).sections.map((section) => section.id)).not.toContain('errorBars');
	});

	it('edits the amount, direction, caps and custom values of one series', () => {
		let data = columnChart();
		const set = (field: string, value: string | number | boolean) => {
			const next = applyChartFormatField(data, 'series:0', field, value);
			expect(next, field).not.toBeNull();
			data = next!;
			return data.series[0]!.errBars;
		};
		expect(set('errorBars.type', 'cust')).toEqual([
			{
				direction: 'y',
				barType: 'both',
				valType: 'cust',
				customPlus: [1, 1, 1],
				customMinus: [1, 1, 1],
			},
		]);
		const section = seriesOptions(data).sections.find((item) => item.id === 'errorBars')!;
		expect(section.fields.map((field) => field.id)).toEqual([
			'errorBars.type',
			'errorBars.side',
			'errorBars.cap',
			'errorBars.plusValues',
			'errorBars.minusValues',
		]);
		expect(section.note?.key).toBe('chartFormat.errorBars.customValuesHint');
		expect(set('errorBars.plusValues', '2, 3; 4')?.[0]?.customPlus).toEqual([2, 3, 4]);
		expect(set('errorBars.minusValues', '0.5')?.[0]?.customMinus).toEqual([0.5, 0.5, 0.5]);
		// PowerPoint keeps custom amounts as values: a range and negative values are refused.
		expect(
			applyChartFormatField(data, 'series:0', 'errorBars.plusValues', 'Sheet1!$C$2:$C$4'),
		).toBeNull();
		expect(applyChartFormatField(data, 'series:0', 'errorBars.plusValues', '-1')).toBeNull();
		expect(set('errorBars.side', 'plus')?.[0]?.barType).toBe('plus');
		expect(set('errorBars.cap', false)?.[0]?.noEndCap).toBe(true);
		expect(set('errorBars.type', 'percentage')).toEqual([
			{ direction: 'y', barType: 'plus', valType: 'percentage', noEndCap: true },
		]);
		expect(set('errorBars.value', 10)?.[0]?.val).toBe(10);
		// The other series is untouched.
		expect(data.series[1]!.errBars).toBeUndefined();
		expect(set('errorBars.type', 'none')).toEqual([]);
	});

	it('keeps X error bars and refuses ambiguous Y error bars', () => {
		const data = columnChart();
		data.series[0]!.errBars = [{ direction: 'x', barType: 'both', valType: 'fixedVal', val: 2 }];
		const next = applyChartFormatField(data, 'series:0', 'errorBars.type', 'stdErr')!;
		expect(next.series[0]!.errBars).toEqual([
			{ direction: 'x', barType: 'both', valType: 'fixedVal', val: 2 },
			{ direction: 'y', barType: 'both', valType: 'stdErr', val: 1 },
		]);
		const twice = columnChart();
		twice.series[0]!.errBars = [
			{ direction: 'y', barType: 'both', valType: 'fixedVal', val: 2 },
			{ direction: 'y', barType: 'plus', valType: 'percentage', val: 5 },
		];
		expect(applyChartFormatField(twice, 'series:0', 'errorBars.cap', false)).toBeNull();
		const section = seriesOptions(twice).sections.find((item) => item.id === 'errorBars')!;
		expect(section.note?.key).toBe('chartFormat.analysis.multipleErrorBars');
	});

	it('has an English caption for every Error Bars string', () => {
		let data = applyChartFormatField(columnChart(), 'series:0', 'errorBars.type', 'cust')!;
		const keys = new Set<string>();
		for (const model of [
			seriesOptions(data),
			seriesOptions(
				(data = applyChartFormatField(data, 'series:0', 'errorBars.type', 'fixedVal')!),
			),
		]) {
			for (const section of model.sections) {
				if (section.id !== 'errorBars') continue;
				for (const text of [section.title, section.note])
					if (text) keys.add(pptxChartFormatKey(text.key));
				for (const field of section.fields) {
					keys.add(pptxChartFormatKey(field.label.key));
					for (const choice of 'choices' in field ? field.choices : [])
						keys.add(pptxChartFormatKey(choice.label.key));
				}
			}
		}
		expect([...keys].filter((key) => !(key in translationsEn))).toEqual([]);
	});

	it('adds, edits and removes several trendlines of one series', () => {
		let data = columnChart();
		const set = (field: string, value: string | number | boolean) => {
			const next = applyChartFormatField(data, 'series:0', field, value);
			expect(next, field).not.toBeNull();
			data = next!;
			return data.series[0]!.trendlines;
		};
		expect(set('trendline.type', 'poly')).toEqual([{ trendlineType: 'polynomial', order: 2 }]);
		expect(set('trendline.order', 3)?.[0]?.order).toBe(3);
		expect(set('trendline.equation', true)?.[0]?.displayEq).toBe(true);
		expect(set('trendline.add', 'movingAvg')).toEqual([
			{ trendlineType: 'polynomial', order: 3, displayEq: true },
			{ trendlineType: 'movingAvg', period: 2 },
		]);
		expect(set('trendline.1.period', 3)?.[1]?.period).toBe(3);
		expect(set('trendline.1.forward', 1)?.[1]?.forward).toBe(1);
		const ids = seriesOptions(data).sections.map((section) => section.id);
		expect(ids).toEqual(expect.arrayContaining(['trendlines', 'trendline-1', 'errorBars']));
		// A refused option leaves the chart as it is.
		expect(applyChartFormatField(data, 'series:0', 'trendline.order', 9)).toBeNull();
		expect(set('trendline.type', 'none')).toEqual([
			{ trendlineType: 'movingAvg', period: 3, forward: 1 },
		]);
		expect(data.series[1]!.trendlines).toBeUndefined();
	});
});
