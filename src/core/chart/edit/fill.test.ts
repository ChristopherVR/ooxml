import { describe, expect, it } from 'vitest';
import { chartPaletteSeriesColorChoice, findChartColorPalette } from '../color-palettes';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { chartSpaceSeries } from './draft';
import {
	editChartElementGradient,
	editChartSeriesGradient,
	setChartDataPointFill,
	setChartElementFill,
	setChartSeriesColor,
	setChartSeriesTransparency,
} from './fill';
import { applyChartColorPalette, setChartStyle } from './style';
import { setChartType } from './type';
import { columnChart } from './test-support/column-chart';

const red = { kind: 'srgb' as const, value: 'FF0000', transforms: [] };

describe('fill edits', () => {
	it('colours a series fill, keeping its transparency and patching the kept spPr', () => {
		const translucent = setChartSeriesTransparency(columnChart(), 0, 25);
		expect(chartSpaceSeries(translucent)[0]!.spPr?.fill).toEqual({
			kind: 'solid',
			color: { kind: 'srgb', value: '4472C4', transforms: [{ name: 'alpha', value: '75000' }] },
		});
		const next = setChartSeriesColor(translucent, 0, red);
		const xml = writeChartSpace(next);
		expect(xml).toContain(
			'<c:spPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>',
		);
		expect(parseChartSpace(xml).issues).toEqual([]);
	});

	it('colours a line series by its line', () => {
		const line = setChartType(columnChart(), { kind: 'line' });
		const next = setChartSeriesColor(line, 1, red);
		const shape = chartSpaceSeries(next)[1]!.spPr!;
		expect(shape.line?.fill).toEqual({ kind: 'solid', color: red });
		expect(shape.fill).toEqual(chartSpaceSeries(line)[1]!.spPr!.fill);
	});

	it('adds and removes a data point override', () => {
		const next = setChartDataPointFill(columnChart(), 0, 2, { kind: 'solid', color: red });
		expect(chartSpaceSeries(next)[0]!.dataPoints).toEqual([
			{ index: 2, spPr: { fill: { kind: 'solid', color: red } } },
		]);
		expect(parseChartSpace(writeChartSpace(next)).issues).toEqual([]);
		expect(chartSpaceSeries(setChartDataPointFill(next, 0, 2, undefined))[0]!.dataPoints).toEqual(
			[],
		);
	});

	it('fills chart elements and edits their gradients', () => {
		const next = setChartElementFill(columnChart(), 'plotArea', { kind: 'solid', color: red });
		expect(next.plotArea.spPr?.fill).toEqual({ kind: 'solid', color: red });
		const gradient = editChartElementGradient(next, 'plotArea', { kind: 'create' })!;
		expect(gradient.space.plotArea.spPr?.fill).toMatchObject({
			kind: 'gradient',
			stops: [{ position: 0, color: red }, { position: 100 }],
		});
		expect(() => setChartElementFill(setChartTypeNoTitle(), 'title', { kind: 'none' })).toThrow(
			RangeError,
		);
		const series = editChartSeriesGradient(columnChart(), 0, { kind: 'create' }, red)!;
		expect(chartSpaceSeries(series.space)[0]!.spPr?.fill).toMatchObject({ kind: 'gradient' });
		expect(editChartSeriesGradient(columnChart(), 0, { kind: 'angle', value: 45 }, red)).toBe(
			undefined,
		);
	});
});

const setChartTypeNoTitle = () => {
	const space = columnChart();
	delete space.title;
	return space;
};

describe('style edits', () => {
	it('writes modern styles with their fallback', () => {
		expect(setChartStyle(columnChart(), 102)).toMatchObject({ c14Style: 102, style: 2 });
		expect(setChartStyle(setChartStyle(columnChart(), 102), 7).c14Style).toBeUndefined();
		expect(() => setChartStyle(columnChart(), 60)).toThrow(RangeError);
	});

	it('recolours palette colours and keeps custom ones', () => {
		const palette10 = findChartColorPalette(10)!;
		const space = setChartSeriesColor(
			columnChart(),
			0,
			chartPaletteSeriesColorChoice(palette10, 0, 2),
		);
		const next = applyChartColorPalette(space, 10, 11);
		const series = chartSpaceSeries(next);
		expect(series[0]!.spPr?.fill).toEqual({
			kind: 'solid',
			color: chartPaletteSeriesColorChoice(findChartColorPalette(11)!, 0, 2),
		});
		expect(series[1]!.spPr?.fill).toEqual(chartSpaceSeries(space)[1]!.spPr?.fill);
	});
});
