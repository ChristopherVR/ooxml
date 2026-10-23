import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../types';
import { planChartExNeutralSave, saveChartExPartThroughNeutral } from './chart-ex-neutral-save';

const PART =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$3</cx:f><cx:lvl ptCount="2"><cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Old</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel" uniqueId="{U}"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Sales</cx:v></cx:txData></cx:tx><cx:dataId val="0" /></cx:series></cx:plotAreaRegion></cx:plotArea><cx:legend pos="t" align="ctr" overlay="0" /></cx:chart></cx:chartSpace>';

const BEFORE: PptxChartData = {
	chartType: 'funnel',
	title: 'Old',
	categories: ['A', 'B'],
	categoryLevels: [['A', 'B']],
	series: [{ name: 'Sales', values: [1, 2] }],
	style: { hasTitle: true, hasLegend: true, legendPosition: 't' },
};

const edited = (recipe: (chart: PptxChartData) => void): PptxChartData => {
	const chart = structuredClone(BEFORE);
	recipe(chart);
	return chart;
};

describe('saveChartExPartThroughNeutral', () => {
	it('saves claimed edits and keeps every other byte (formulas, ids, empty-element form)', () => {
		const after = edited((chart) => {
			chart.series[0]!.name = 'Revenue';
			chart.series[0]!.values = [5, 6];
			chart.style!.hasLegend = false;
		});
		expect(saveChartExPartThroughNeutral(PART, BEFORE, after)).toBe(
			PART.replace('<cx:v>Sales</cx:v>', '<cx:v>Revenue</cx:v>')
				.replace('<cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt>', '')
				.replace(
					'<cx:lvl ptCount="2"></cx:lvl></cx:numDim>',
					'<cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">5</cx:pt><cx:pt idx="1">6</cx:pt></cx:lvl></cx:numDim>',
				)
				.replace('<cx:legend pos="t" align="ctr" overlay="0" />', ''),
		);
	});

	it('writes a new title over a cell-linked one and removes a hidden one', () => {
		const titled = saveChartExPartThroughNeutral(
			PART,
			BEFORE,
			edited((chart) => {
				chart.title = 'New';
			}),
		);
		expect(titled).toContain(
			'<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr /><a:p><a:r><a:t>New</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>',
		);
		const hidden = saveChartExPartThroughNeutral(
			PART,
			BEFORE,
			edited((chart) => {
				chart.style!.hasTitle = false;
			}),
		);
		expect(hidden).not.toContain('<cx:title');
	});

	it('leaves unclaimed edits and unreadable parts to the object-tree update', () => {
		const coloured = edited((chart) => {
			chart.series[0]!.color = '#FF0000';
		});
		expect(planChartExNeutralSave(BEFORE, coloured)).toBeUndefined();
		expect(saveChartExPartThroughNeutral(PART, BEFORE, coloured)).toBeUndefined();
		const added = edited((chart) => {
			chart.series.push({ name: 'More', values: [1] });
		});
		expect(planChartExNeutralSave(BEFORE, added)).toBeUndefined();
		const renamed = edited((chart) => {
			chart.series[0]!.name = 'X';
		});
		expect(
			saveChartExPartThroughNeutral(
				PART.replace('</cx:series>', '<cx:x/></cx:series>'),
				BEFORE,
				renamed,
			),
		).toBeUndefined();
		expect(planChartExNeutralSave(BEFORE, structuredClone(BEFORE))).toEqual([]);
	});
});
