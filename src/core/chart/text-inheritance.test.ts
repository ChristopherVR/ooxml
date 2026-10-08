import { expect, it } from 'vitest';
import native from './excel-chart-text-inheritance.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { chartXml } from '../xlsx/write/chart';
import { readChartStyle } from './read-style';
import { NS, parseXml, first, buildXml } from '../xml';
import { renderChartSvg } from './render/chart-svg';

for (const sample of native.cases)
	it(`matches native chart-space text inheritance for ${sample.referenceName}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const style = sample.parts['xl/charts/style1.xml' as keyof typeof sample.parts];
		if (style) chart.styleDefinition = readChartStyle(style)!;
		const book = createWorkbook();
		book.theme.minorFont = 'Aptos Narrow';
		book.theme.majorFont = 'Aptos Display';
		const verify = (target: typeof chart) => {
			const view = chartView(book, 0, target, () => []);
			for (const [part, font] of [
				['title', sample.titleText],
				['legend', sample.legendText],
				['categoryAxis', sample.categoryText],
				['valueAxis', sample.valueText],
			] as const) {
				const face = font.name.startsWith('+mn-')
					? book.theme.minorFont
					: font.name.startsWith('+mj-')
						? book.theme.majorFont
						: font.name;
				expect(view.appearance?.[part]).toMatchObject({
					fontSize: font.size,
					bold: font.bold,
					italic: font.italic,
					color: font.color,
					typeface: face,
				});
			}
		};
		verify(chart);
		const heading = [
			...parseXml(
				renderChartSvg(
					chartView(book, 0, chart, () => []),
					600,
					400,
				),
			).getElementsByTagName('text'),
		].find((node) => node.textContent === 'Native style')!;
		expect(Number(heading.getAttribute('font-size'))).toBeCloseTo((sample.titleText.size! * 4) / 3);
		expect(heading.getAttribute('font-style') === 'italic').toBe(sample.titleText.italic);
		const regeneratedXml = chartXml({ ...chart, chartType: 'line', title: 'Edited native title' });
		const regenerated = parseChart(regeneratedXml, chart.anchor, '');
		if (chart.styleDefinition) regenerated.styleDefinition = chart.styleDefinition;
		verify(regenerated);
		expect(regenerated.title).toBe('Edited native title');
		const rootText = (xml: string) => buildXml(first(parseXml(xml).documentElement, 'txPr', NS.c)!);
		expect(rootText(regeneratedXml)).toBe(rootText(source));
	});

it('retains chart-wide text on a minimal regenerated chart without fill or axis patches', () => {
	const source = `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}"><c:chart><c:plotArea/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="2000"><a:latin typeface="Arial"/></a:defRPr></a:pPr></a:p></c:txPr></c:chartSpace>`;
	const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
	const xml = chartXml({ ...chart, chartType: 'line', title: 'New title' });
	const reloaded = parseChart(xml, chart.anchor, '');
	expect(reloaded.formatting?.textDefaults).toMatchObject({ fontSize: 20, typeface: 'Arial' });
	expect(chartView(createWorkbook(), 0, reloaded, () => []).appearance?.title).toMatchObject({
		fontSize: 24,
		typeface: 'Arial',
	});
});
