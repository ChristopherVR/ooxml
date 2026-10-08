import { expect, it } from 'vitest';
import { NS, parseXml } from '../xml';
import native from './excel-built-in-text-styles.json';
import { readChartFormatting } from './read-formatting';
import { builtInChartStyleXml, readBuiltInChartStyle } from './built-in-text-style';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { renderChartSvg } from './render/chart-svg';
import { chartXml } from '../xlsx/write/chart';
import { readChartStyle } from './read-style';
import external from './excel-chart-styles.json';

for (const sample of native.cases)
	it(`matches isolated native built-in text style ${sample.requestedStyle}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		expect(chart.formatting?.builtInStyle).toBe(sample.requestedStyle);
		const book = createWorkbook();
		book.theme.minorFont = sample.titleText.name;
		const view = chartView(book, 0, chart, () => []);
		for (const [part, font] of [
			['title', sample.titleText],
			['legend', sample.legendText],
			['categoryAxis', sample.categoryText],
			['valueAxis', sample.valueText],
		] as const) {
			expect(view.appearance?.[part]).toMatchObject({
				fontSize: font.size,
				bold: font.bold,
				italic: font.italic,
				color: font.color,
				typeface: font.name,
			});
		}
		const svg = parseXml(renderChartSvg(view, 600, 400));
		const heading = [...svg.getElementsByTagName('text')].find(
			(node) => node.textContent === 'Native style',
		)!;
		expect(heading.getAttribute('font-size')).toBe(String((sample.titleText.size! * 4) / 3));
		expect(heading.getAttribute('font-weight')).toBe('bold');
		const regenerated = parseChart(chartXml({ ...chart, chartType: 'line' }), chart.anchor, '');
		expect(regenerated.formatting?.builtInStyle).toBe(sample.requestedStyle);
		expect(chartView(book, 0, regenerated, () => []).appearance?.title).toMatchObject({
			fontSize: sample.titleText.size,
			bold: sample.titleText.bold,
			italic: sample.titleText.italic,
			color: sample.titleText.color,
			typeface: sample.titleText.name,
		});
	});

it('checks required namespaces and keeps unsupported styles distinct from absent defaults', () => {
	const xml = (content: string) =>
		parseXml(`<c:chartSpace xmlns:c="${NS.c}">${content}<c:chart/></c:chartSpace>`).documentElement;
	expect(readBuiltInChartStyle(xml(''))).toBe(2);
	expect(readBuiltInChartStyle(xml('<c:style val="999"/>'))).toBe(999);
	expect(readBuiltInChartStyle(xml('<c:style val="invalid"/>'))).toBeUndefined();
	expect(readBuiltInChartStyle(xml(builtInChartStyleXml(102)))).toBe(102);
	const unknown = builtInChartStyleXml(102).replace('Requires="c14"', 'Requires="c14 other"');
	expect(readBuiltInChartStyle(xml(unknown))).toBe(2);
	expect(builtInChartStyleXml(999)).toBe('');
	expect(readChartFormatting(xml('<c:style val="41"/>'))?.builtInStyle).toBe(41);
});

it('follows theme changes and allows external and direct text formatting to override built-in defaults', () => {
	const sample = native.cases.find((item) => item.requestedStyle === 102)!;
	const chart = parseChart(
		sample.parts['xl/charts/chart1.xml'],
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'',
	);
	const book = createWorkbook();
	book.theme.minorFont = 'Changed Minor';
	book.theme.colors[1] = '123456';
	expect(chartView(book, 0, chart, () => []).appearance?.title).toMatchObject({
		typeface: 'Changed Minor',
		color: '#123456',
		fontSize: 18,
		bold: true,
	});
	chart.styleDefinition = readChartStyle(external.cases[0]!.parts['xl/charts/style1.xml'])!;
	expect(chartView(book, 0, chart, () => []).appearance?.title).toMatchObject({
		fontSize: 14,
		bold: false,
	});
	chart.formatting!.entries.title = {
		...chart.formatting!.entries.title!,
		fontSize: 24,
		bold: true,
		typeface: 'Explicit Face',
		textColor: { kind: 'srgb', value: 'ABCDEF', transforms: [] },
	};
	expect(chartView(book, 0, chart, () => []).appearance?.title).toMatchObject({
		fontSize: 24,
		bold: true,
		typeface: 'Explicit Face',
		color: '#ABCDEF',
	});
});

it('uses modern defaults for newly authored parts while retaining imported legacy choices', () => {
	const chart = parseChart(
		native.cases[0]!.parts['xl/charts/chart1.xml'],
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'',
	);
	const { formatting: _formatting, ...fresh } = chart;
	const parsed = parseChart(chartXml(fresh), chart.anchor, '');
	expect(parsed.formatting?.builtInStyle).toBe(102);
	expect(chartView(createWorkbook(), 0, parsed, () => []).appearance?.title?.fontSize).toBe(18);
	expect(parseChart(chartXml(chart), chart.anchor, '').formatting?.builtInStyle).toBe(1);
});
