import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../../chart/parse-space';
import { NS } from '../../xml/index';
import type { ChartObject } from '../model';
import { parseChart } from '../read/chart';
import { chartXml } from './chart';

const anchor = { from: { col: 0, row: 0 }, to: { col: 5, row: 10 } } as ChartObject['anchor'];

const column: ChartObject = {
	kind: 'chart',
	anchor,
	chartType: 'column',
	title: 'Sales',
	showLegend: true,
	series: [
		{
			name: 'North',
			nameRef: 'Data!$B$1',
			categoriesRef: 'Data!$A$2:$A$3',
			valuesRef: 'Data!$B$2:$B$3',
			categories: ['Q1', 'Q2'],
			values: [10, null],
			pointColors: { 1: { kind: 'srgb', value: 'FF0000', transforms: [] } },
		},
	],
};

describe('chartXml (new chart parts through the shared chart writer)', () => {
	it('writes the part Excel accepts, declaring each namespace once on the root', () => {
		expect(chartXml(column)).toBe(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
				`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/>` +
				`<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice Requires="c14" xmlns:c14="${NS.c14}"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>` +
				'<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
				'<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
				'<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>North</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
				'<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' +
				'<c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt>' +
				'<c:cat><c:strRef><c:f>Data!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>' +
				'<c:val><c:numRef><c:f>Data!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>' +
				'<c:gapWidth val="150"/><c:axId val="500000001"/><c:axId val="500000002"/></c:barChart>' +
				'<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
				'<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>' +
				'</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>',
		);
	});

	it('reads back as the chart it was written from, for every generated family', () => {
		for (const chartType of [
			'column',
			'bar',
			'line',
			'area',
			'pie',
			'doughnut',
			'scatter',
			'radar',
		] as const) {
			const chart: ChartObject = { ...column, chartType, legendPosition: 'b' };
			const xml = chartXml(chart);
			expect(parseChartSpace(xml).issues, chartType).toEqual([]);
			const read = parseChart(xml, anchor, 'xl/charts/chart1.xml');
			expect(read, chartType).toMatchObject({
				chartType,
				title: 'Sales',
				showLegend: true,
				legendPosition: 'b',
				series: [
					{ name: 'North', nameRef: 'Data!$B$1', valuesRef: 'Data!$B$2:$B$3', values: [10, null] },
				],
			});
		}
	});
});
