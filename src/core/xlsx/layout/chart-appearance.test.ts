import { expect, it } from 'vitest';
import { readChartStyle } from '../../chart/read-style';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { chartView } from './chart-view';
import { renderChartSvg } from './chart-svg';
import { NS, parseXml } from '../../xml/index';

it('prefers direct run formatting and no-fill/zero-width overrides while following theme changes', () => {
	const chart = parseChart(
		`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart><c:title><c:txPr><a:p><a:pPr><a:defRPr sz="1800" b="1"/></a:pPr></a:p></c:txPr><c:tx><c:rich><a:p><a:pPr><a:defRPr i="1"/></a:pPr><a:r><a:rPr sz="2400" b="0"><a:latin typeface="+mj-lt"/><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:rPr><a:t>Custom title</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart/><c:catAx><c:tickLblPos val="none"/></c:catAx></c:plotArea></c:chart><c:spPr><a:noFill/><a:ln w="0"><a:noFill/></a:ln></c:spPr></c:chartSpace>`,
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'xl/charts/chart1.xml',
	);
	chart.styleDefinition = readChartStyle(
		`<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="${NS.a}"><cs:title><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000"/></cs:title><cs:chartArea><cs:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></cs:spPr></cs:chartArea></cs:chartStyle>`,
	)!;
	const book = createWorkbook();
	book.theme.majorFont = 'Custom Heading';
	book.theme.colors[5] = '123456';
	const view = chartView(book, 0, chart, () => []);
	expect(view.appearance?.title).toMatchObject({
		fontSize: 24,
		bold: false,
		italic: true,
		color: '#123456',
		typeface: 'Custom Heading',
	});
	expect(view.appearance?.chartArea).toMatchObject({
		fillColor: 'none',
		lineColor: 'none',
		lineWidth: 0,
	});
	expect(view.appearance?.categoryAxis?.labelsVisible).toBe(false);
	book.theme.majorFont = 'Updated Heading';
	book.theme.colors[5] = '654321';
	expect(chartView(book, 0, chart, () => []).appearance?.title).toMatchObject({
		color: '#654321',
		typeface: 'Updated Heading',
	});
	const svg = parseXml(renderChartSvg(view, 640, 400));
	const title = [...svg.getElementsByTagName('text')].find(
		(text) => text.textContent === 'Custom title',
	)!;
	expect(title.getAttribute('font-size')).toBe('32');
	expect(title.getAttribute('font-weight')).toBe('normal');
	expect(title.getAttribute('font-style')).toBe('italic');
});

it('gives adjacent chart gradients distinct targets without changing the view model', () => {
	const chart = {
		kind: 'chart' as const,
		chartType: 'column' as const,
		anchor: { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		series: [],
		showLegend: false,
		formatting: {
			sourceXml: '',
			entries: {
				chartArea: {
					sourceXml: '',
					fill: {
						kind: 'gradient' as const,
						angle: 90,
						stops: [
							{ position: 0, color: { kind: 'srgb' as const, value: 'FF0000', transforms: [] } },
							{ position: 100, color: { kind: 'srgb' as const, value: '0000FF', transforms: [] } },
						],
					},
				},
			},
		},
	};
	const view = chartView(createWorkbook(), 0, chart, () => []);
	const before = structuredClone(view);
	const a = parseXml(renderChartSvg(view, 640, 400));
	const b = parseXml(renderChartSvg(view, 640, 400));
	expect(a.getElementsByTagName('linearGradient')[0]!.getAttribute('id')).not.toBe(
		b.getElementsByTagName('linearGradient')[0]!.getAttribute('id'),
	);
	expect(view).toEqual(before);
});
