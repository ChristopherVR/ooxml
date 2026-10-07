import { expect, it } from 'vitest';
import native from './excel-chart-title-text.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { renderChartSvg } from '../xlsx/layout/chart-svg';
import { chartXml } from '../xlsx/write/chart';
import { NS, parseXml, first, buildXml } from '../xml';

const richXml = (source: string) => {
	const title = first(first(parseXml(source).documentElement, 'chart', NS.c), 'title', NS.c);
	return buildXml(first(first(title, 'tx', NS.c), 'rich', NS.c)!);
};

for (const sample of native.cases)
	it(`renders and preserves native mixed title ${sample.referenceName}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const book = createWorkbook();
		book.theme.minorFont = 'Aptos Narrow';
		const verify = (target: typeof chart) => {
			const view = chartView(book, 0, target, () => []);
			const characters = view.titleText!.paragraphs.flatMap((paragraph) =>
				paragraph.runs.flatMap((run) =>
					[...run.text]
						.filter((text) => text !== '\n')
						.map((text) => ({ text, font: run.appearance })),
				),
			);
			const expected = sample.titleCharacters.filter((item) => !['\n', '\r'].includes(item.text));
			expect(characters.map((item) => item.text).join('')).toBe(
				expected.map((item) => item.text).join(''),
			);
			for (const [index, actual] of characters.entries()) {
				const font = expected[index]!.font;
				expect(actual.font).toMatchObject({
					fontSize: font.size,
					bold: font.bold,
					italic: font.italic,
					underline: font.underline,
					color: font.color,
					typeface: font.name,
				});
			}
			const svg = parseXml(renderChartSvg(view, 640, 400));
			const spans = [...svg.getElementsByTagName('tspan')];
			expect(spans.map((node) => node.textContent).join('')).toBe(
				expected.map((item) => item.text).join(''),
			);
			expect(spans[0]!.getAttribute('fill')).toBe('#FF0000');
			expect(spans[0]!.getAttribute('font-size')).toBe('32');
			expect(spans.find((node) => node.textContent === 'growth')!.getAttribute('font-style')).toBe(
				'italic',
			);
			if (!sample.referenceName.endsWith('single')) {
				if (sample.referenceName.endsWith('paragraphs'))
					expect(view.titleText!.paragraphs).toHaveLength(2);
				const forecast = spans.find((node) => node.textContent === 'Forecast')!;
				expect(forecast.getAttribute('text-decoration')).toBe('underline');
				expect(Number(forecast.getAttribute('y'))).toBeGreaterThan(
					Number(spans[0]!.getAttribute('y')),
				);
			}
		};
		verify(chart);
		const regenerated = chartXml({ ...chart, chartType: 'line' });
		expect(richXml(regenerated)).toBe(richXml(source));
		verify(parseChart(regenerated, chart.anchor, ''));
		const edited = parseChart(
			chartXml({ ...chart, chartType: 'line', title: 'Edited title' }),
			chart.anchor,
			'',
		);
		expect(edited.title).toBe('Edited title');
		expect(chartView(book, 0, edited, () => []).titleText).toBeUndefined();
	});
