import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../types';
import { planNeutralChartSave } from './chart-neutral-save-plan';
import { saveChartPartThroughNeutral } from './chart-neutral-save';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const OF_PIE =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="${C}"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
	'<c:ofPieChart><c:ofPieType val="pie"/><c:varyColors val="1"/>' +
	'<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="4"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt><c:pt idx="3"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser>' +
	'<c:gapWidth val="100"/><c:splitType val="auto"/><c:secondPieSize val="75"/><c:serLines/></c:ofPieChart>' +
	'</c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

const before = (): PptxChartData =>
	({
		chartType: 'ofPie',
		categories: [],
		series: [{ name: '', idx: 0, values: [4, 3, 2, 1] }],
		ofPieOptions: {
			ofPieType: 'pie',
			splitType: 'auto',
			secondPieSize: 75,
			serLines: true,
			gapWidth: 100,
		},
	}) as PptxChartData;

const edited = (edit: (chart: PptxChartData) => void) => {
	const chart = before();
	edit(chart);
	return chart;
};

describe('neutral save of chart group subtypes', () => {
	it('writes a pie of pie split, plot gap and series lines in schema order', () => {
		const saved = saveChartPartThroughNeutral(
			OF_PIE,
			before(),
			edited((chart) => {
				chart.ofPieOptions = {
					ofPieType: 'bar',
					splitType: 'cust',
					custSplit: [2, 3],
					secondPieSize: 40,
					serLines: false,
					gapWidth: 150,
				};
			}),
		);
		expect(saved).toBeDefined();
		expect(saved!.xml).toContain('<c:ofPieType val="bar"/><c:varyColors val="1"/>');
		expect(saved!.xml).toContain(
			'<c:gapWidth val="150"/><c:splitType val="cust"/><c:custSplit><c:secondPiePt val="2"/><c:secondPiePt val="3"/></c:custSplit><c:secondPieSize val="40"/></c:ofPieChart>',
		);
	});

	it('leaves cleared settings the object-tree writers keep to them', () => {
		// Cleared pie of pie options, a cleared bubble option and absent series lines keep the
		// authored elements on the object-tree path, so the neutral plan does not claim them.
		expect(
			planNeutralChartSave(
				before(),
				edited((chart) => delete chart.ofPieOptions),
			),
		).toBe(undefined);
		expect(
			planNeutralChartSave(
				before(),
				edited((chart) => {
					chart.ofPieOptions = { ofPieType: 'pie', splitType: 'auto', secondPieSize: 75 };
				}),
			),
		).toBe(undefined);
		const bubble = { ...before(), chartType: 'bubble', bubbleOptions: { bubbleScale: 50 } };
		expect(
			planNeutralChartSave(
				bubble as PptxChartData,
				{
					...bubble,
					bubbleOptions: {},
				} as PptxChartData,
			),
		).toBe(undefined);
	});

	it('leaves a subtype setting the chart type does not write to the object-tree path', () => {
		const bar = { ...before(), chartType: 'bar' } as PptxChartData;
		delete bar.ofPieOptions;
		expect(planNeutralChartSave(bar, { ...bar, radarStyle: 'filled' })).toBe(undefined);
		expect(planNeutralChartSave(bar, { ...bar, barShape: 'cone' })).toBe(undefined);
	});

	it('falls back when the shared operation refuses a value', () => {
		const saved = saveChartPartThroughNeutral(
			OF_PIE,
			before(),
			edited((chart) => {
				chart.ofPieOptions = { ...chart.ofPieOptions!, secondPieSize: 300 };
			}),
		);
		expect(saved).toBeUndefined();
	});
});
