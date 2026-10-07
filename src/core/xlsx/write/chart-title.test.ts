import { expect, it } from 'vitest';
import { NS, parseXml, first } from '../../xml';
import type { ChartObject } from '../model';
import { parseChart } from '../read/chart';
import { chartXml } from './chart';
import { patchChartPart } from './chart-patch';

const chart: ChartObject = {
	kind: 'chart',
	chartType: 'column',
	showLegend: false,
	anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	series: [],
};

it.each(['new', 'insert', 'reference'] as const)(
	'retains native rich property nodes when writing a %s title',
	(mode) => {
		const title = 'Revenue < & expenses';
		let xml: string;
		if (mode === 'new') xml = chartXml({ ...chart, title });
		else {
			let source = chartXml(chart);
			if (mode === 'reference')
				source = source.replace(
					'<c:chart>',
					'<c:chart><c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Old</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>',
				);
			xml = patchChartPart(source, { ...parseChart(source, chart.anchor, ''), title })!;
		}
		const doc = parseXml(xml);
		const parent = first(first(doc.documentElement, 'chart', NS.c), 'title', NS.c);
		const paragraph = first(first(first(parent, 'tx', NS.c), 'rich', NS.c), 'p', NS.a);
		expect(first(first(paragraph, 'pPr', NS.a), 'defRPr', NS.a)).toBeDefined();
		expect(first(first(paragraph, 'r', NS.a), 'rPr', NS.a)).toBeDefined();
		expect(parseChart(xml, chart.anchor, '').title).toBe(title);
	},
);
