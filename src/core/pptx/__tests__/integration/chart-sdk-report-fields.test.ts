/**
 * The report fields of `ChartInput` (label anchor, axis and gridline lines with dashes, tick
 * marks, crossing, title runs), each saved through the shared chart core and reloaded: the part
 * holds them, the reloaded model reads them back, an omitted field writes nothing, and an edge
 * value is written or refused, never dropped.
 */
import { describe, expect, it } from 'vitest';

import { createChartElement } from '../../core/builders/sdk/ElementFactory';
import type { ChartInput } from '../../core/builders/sdk/types';
import type { PptxChartAxisFormatting, PptxChartType } from '../../core/types';
import { firstElement, newChartDeck } from '../helpers/new-chart-deck';

const CATEGORIES = ['a', 'b', 'c'];
const SERIES: ChartInput['series'] = [{ name: 'A', values: [0.2, 0.5, 0.9] }];

async function saved(input: Partial<ChartInput>, type: PptxChartType = 'bar') {
	const { save, element } = await newChartDeck(type, {
		categories: CATEGORIES,
		series: SERIES,
		...input,
	});
	const result = await save();
	// Nothing falls back to the legacy generator, which would drop fields.
	expect(result.warnings).toStrictEqual([]);
	return { ...result, element };
}

const axisOf = (axes: PptxChartAxisFormatting[] | undefined, type: 'catAx' | 'valAx') =>
	axes?.find((axis) => axis.axisType === type);

describe('report fields of ChartInput', () => {
	it('writes nothing for omitted ticks, crossing, gridlines and anchor', async () => {
		const { xml, chart } = await saved({});
		const value = firstElement(xml, 'c:valAx');
		for (const name of ['majorTickMark', 'minorTickMark', 'crosses', 'crossesAt', 'crossBetween'])
			expect(value).not.toContain(`<c:${name} `);
		expect(xml).not.toContain('Gridlines>');
		expect(xml).not.toContain('anchor=');
		const axis = axisOf(chart.axes, 'valAx');
		expect(axis?.majorTickMark).toBeUndefined();
		expect(axis?.crosses).toBeUndefined();
		expect(axis?.majorGridlines).toBeFalsy();
	});

	it('reloads a label anchor, rotation, number format and no wrap', async () => {
		const { xml, chart } = await saved({
			series: [
				{
					...SERIES[0]!,
					dataLabels: {
						showValue: true,
						position: 'ctr',
						numberFormat: '0%',
						rotation: -90,
						anchor: 'ctr',
						wrap: false,
					},
				},
			],
		});
		const bodyPr = /<a:bodyPr[^>]*>/u.exec(firstElement(xml, 'c:dLbls'))![0];
		expect(bodyPr).toContain('anchor="ctr"');
		expect(bodyPr).toContain('wrap="none"');
		expect(bodyPr).toContain('rot="-5400000"');
		expect(chart.series[0]!.dataLabelOptions).toMatchObject({
			showValue: true,
			position: 'ctr',
			numberFormat: '0%',
		});
		const formats = Object.entries(chart.elementFormats ?? {}).filter(([key]) =>
			key.startsWith('dataLabels'),
		);
		expect(formats.map(([, format]) => format.text)).toStrictEqual([
			{ rotation: -90, anchor: 'ctr', wrap: false },
		]);
	});

	it('reloads a scaled value axis with dashed minor gridlines and dash-only axis line', async () => {
		const { xml, chart } = await saved({
			axes: {
				valAx: {
					min: 0,
					max: 1,
					majorUnit: 0.25,
					minorUnit: 0.05,
					line: { dash: 'sysDash' },
					majorGridlinesLine: { color: '#ADB5BD', width: 0.75 },
					minorGridlinesLine: { color: '#DEE2E6', width: 0.5, dash: 'dash' },
					majorTickMark: 'none',
					minorTickMark: 'in',
					crosses: 0.5,
					crossBetween: 'midCat',
				},
			},
		});
		const value = firstElement(xml, 'c:valAx');
		expect(firstElement(value, 'c:minorGridlines')).toBe(
			'<c:minorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="DEE2E6"></a:srgbClr></a:solidFill><a:prstDash val="dash"></a:prstDash></a:ln></c:spPr></c:minorGridlines>',
		);
		expect(value).toContain(
			'<c:spPr><a:ln><a:prstDash val="sysDash"></a:prstDash></a:ln></c:spPr>',
		);
		expect(value).toContain('<c:crossesAt val="0.5">');
		expect(axisOf(chart.axes, 'valAx')).toMatchObject({
			min: 0,
			max: 1,
			majorUnit: 0.25,
			minorUnit: 0.05,
			spPr: { strokeDashStyle: 'sysDash' },
			majorGridlines: true,
			majorGridlinesSpPr: { strokeColor: '#ADB5BD', strokeWidth: 0.75 },
			minorGridlines: true,
			minorGridlinesSpPr: { strokeColor: '#DEE2E6', strokeWidth: 0.5, strokeDashStyle: 'dash' },
			majorTickMark: 'none',
			minorTickMark: 'in',
			crossesAt: 0.5,
			crossBetween: 'midCat',
		});
		expect(axisOf(chart.axes, 'valAx')?.crosses).toBeUndefined();
	});

	it("hides gridlines with a 'none' line instead of writing an invisible one", async () => {
		const { xml, chart } = await saved({
			axes: {
				valAx: { majorGridlines: true, majorGridlinesLine: 'none' },
				catAx: { majorGridlinesLine: 'none' },
			},
		});
		expect(firstElement(xml, 'c:catAx')).not.toContain('Gridlines');
		expect(firstElement(firstElement(xml, 'c:valAx'), 'c:majorGridlines')).toContain(
			'<a:ln><a:noFill>',
		);
		expect(axisOf(chart.axes, 'catAx')?.majorGridlines).toBeFalsy();
		expect(axisOf(chart.axes, 'valAx')).toMatchObject({
			majorGridlines: true,
			majorGridlinesSpPr: { lineNoFill: true },
		});
	});

	it('reloads a reversed category axis with the value axis crossing at max', async () => {
		const { xml, chart } = await saved({
			barDirection: 'bar',
			axes: {
				catAx: { orientation: 'maxMin' },
				valAx: { crosses: 'max', crossBetween: 'between' },
			},
		});
		expect(firstElement(xml, 'c:catAx')).toContain('<c:orientation val="maxMin">');
		expect(firstElement(xml, 'c:valAx')).toContain('<c:crosses val="max">');
		expect(axisOf(chart.axes, 'catAx')?.orientation).toBe('maxMin');
		expect(axisOf(chart.axes, 'valAx')).toMatchObject({ crosses: 'max', crossBetween: 'between' });
	});

	it('reloads title runs with their own fonts and the model title of a loaded chart', async () => {
		const { xml, chart, element } = await saved({
			titleRuns: [
				{ text: 'Factors ', bold: true, size: 16 },
				{ text: 'vs benchmark', italic: true, size: 12, color: '#868E96' },
			],
			titleFont: { family: 'Georgia' },
		});
		expect(firstElement(xml, 'c:title')).toContain('typeface="Georgia"');
		expect(chart.titleRuns).toStrictEqual([
			expect.objectContaining({ text: 'Factors ', bold: true, fontSize: 16 }),
			expect.objectContaining({
				text: 'vs benchmark',
				italic: true,
				fontSize: 12,
				color: '#868E96',
			}),
		]);
		expect(chart.title).toBe('Factors ');
		expect(element.chartData!.title).toBe(chart.title);
	});

	it('refuses edge values instead of dropping them', () => {
		const chart = (input: Partial<ChartInput>) => () =>
			createChartElement('bar', { categories: CATEGORIES, series: SERIES, ...input });
		expect(chart({ axes: { valAx: { majorTickMark: 'outside' as never } } })).toThrow(RangeError);
		expect(chart({ axes: { valAx: { crosses: Number.NaN } } })).toThrow(RangeError);
		expect(chart({ axes: { valAx: { crosses: 'top' as never } } })).toThrow(RangeError);
		expect(chart({ axes: { valAx: { crossBetween: 'on' as never } } })).toThrow(RangeError);
		expect(chart({ axes: { valAx: { line: { dash: 'custom' as never } } } })).toThrow(/preset/u);
		expect(
			chart({
				series: [{ ...SERIES[0]!, dataLabels: { showValue: true, anchor: 'top' as never } }],
			}),
		).toThrow(RangeError);
		expect(chart({ titleRuns: [] })).toThrow(/titleRuns/u);
		expect(chart({ title: 'Other', titleRuns: [{ text: 'Factors' }] })).toThrow(/different/u);
	});
});
