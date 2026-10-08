import { expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read';
import { saveXlsx } from './index';
import { chartXml } from './chart';
import { patchChartPart } from './chart-patch';
import { parseChart } from '../read/chart';
import { parseDrawingFill } from '../../drawingml/drawing-fill';
import { parseXml, NS } from '../../xml';
import type { ChartObject } from '../model';
import type { DrawingFill } from '../../drawingml/types';

const fills: DrawingFill[] = [
	{ kind: 'none' },
	{
		kind: 'solid',
		color: { kind: 'scheme', value: 'accent2', transforms: [{ name: 'alpha', value: '50000' }] },
	},
	...['linear', 'rect', 'circle', 'shape'].map((type) =>
		parseDrawingFill(
			parseXml(
				`<a:spPr xmlns:a="${NS.a}"><a:gradFill rotWithShape="0"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>${type === 'linear' ? '<a:lin ang="5400000" scaled="1"/>' : `<a:path path="${type}"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`}<a:extLst><a:ext uri="retained"/></a:extLst></a:gradFill></a:spPr>`,
			).documentElement,
		)!,
	),
];
const chart = (fill: DrawingFill): ChartObject => ({
	kind: 'chart',
	chartType: 'column',
	showLegend: true,
	title: 'Sales',
	anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
	series: [{ categories: ['A', 'B'], values: [10, 20] }],
	formatting: {
		sourceXml: '',
		entries: Object.fromEntries(
			['chartArea', 'plotArea', 'title', 'legend'].map((part) => [part, { sourceXml: '', fill }]),
		),
	},
});
const modeled = (fill: DrawingFill | undefined) => {
	if (fill?.kind !== 'gradient') return fill;
	const { sourceXml: _sourceXml, ...rest } = fill;
	return rest;
};

it.each(fills.map((fill, index) => ({ fill, index })))(
	'round trips direct chart-element fill $index through repeated saves and regeneration',
	async ({ fill }) => {
		let book = createWorkbook();
		createEditSession(book).addChart(0, chart(fill));
		for (let pass = 0; pass < 2; pass++) {
			book = await loadXlsx(await saveXlsx(book));
			const back = book.sheets[0]!.drawings[0] as ChartObject;
			for (const part of ['chartArea', 'plotArea', 'title', 'legend'])
				expect(modeled(back.formatting?.entries[part]?.fill)).toEqual(modeled(fill));
			back.chartType = 'line';
		}
	},
);

it('patches kept background fills while retaining source styling and untouched XML', () => {
	const model = chart(fills[1]!);
	const source = chartXml(model).replace(
		'</c:spPr>',
		'<a:ln w="12700"/><a:effectLst/><a:extLst><a:ext uri="retained"/></a:extLst></c:spPr>',
	);
	const imported = parseChart(source, model.anchor, 'xl/charts/chart1.xml');
	expect(patchChartPart(source, imported)).toBeUndefined();
	imported.formatting!.entries.chartArea!.fill = fills[5]!;
	const patched = patchChartPart(source, imported)!;
	expect(patched).toContain('uri="retained"');
	expect(patched).toContain('<a:ln w="12700"/>');
	expect(
		modeled(parseChart(patched, model.anchor, '').formatting?.entries.chartArea?.fill),
	).toEqual(modeled(fills[5]));
});

it('writes fills on newly inserted legends and replaced referenced titles', () => {
	const model = chart(fills[1]!);
	model.showLegend = false;
	const source = chartXml(model).replace(
		/<c:title>[\s\S]*?<\/c:title>/,
		`<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Original</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>`,
	);
	const imported = parseChart(source, model.anchor, 'xl/charts/chart1.xml');
	imported.title = 'Replacement';
	imported.showLegend = true;
	for (const part of ['title', 'legend'])
		imported.formatting!.entries[part] = { sourceXml: '', fill: fills[5]! };
	const back = parseChart(patchChartPart(source, imported)!, model.anchor, '');
	for (const part of ['title', 'legend'])
		expect(modeled(back.formatting?.entries[part]?.fill)).toEqual(modeled(fills[5]));
});
