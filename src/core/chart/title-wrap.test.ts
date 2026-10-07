import { expect, it } from 'vitest';
import native from './excel-chart-title-wrap.json';
import browser from './excel-chart-title-wrap-browser-metrics.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { chartRichTitleSvg } from '../xlsx/layout/chart-svg-title-text';
import { chartXml } from '../xlsx/write/chart';
import { NS, first, parseXml, buildXml } from '../xml';

const fonts = browser.fonts as Record<
	string,
	{ ascent: number; descent: number; widths: Record<string, number> }
>;
const richXml = (xml: string) =>
	buildXml(
		first(
			first(first(first(parseXml(xml).documentElement, 'chart', NS.c), 'title', NS.c), 'tx', NS.c),
			'rich',
			NS.c,
		)!,
	);

it('wraps plain titles without truncation while retaining the short-title path', () => {
	const sample = native.cases.find((item) => item.referenceName === 'wrap-long-480')!;
	const chart = parseChart(
		sample.parts['xl/charts/chart1.xml'],
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'',
	);
	const { titleText: _rich, ...plain } = chartView(createWorkbook(), 0, chart, () => []);
	const options = {
		measureText: (text: string) => fonts['normal normal 24px Arial']!.widths[text]!,
	};
	const painted = chartRichTitleSvg(plain, 640, options)!;
	const spans = [...parseXml(`<svg>${painted.markup}</svg>`).getElementsByTagName('tspan')];
	expect(spans).toHaveLength(2);
	expect(spans.map((span) => span.textContent).join(' ')).toBe(plain.title);
	expect(chartRichTitleSvg({ ...plain, title: 'Revenue' }, 640, options)).toBeUndefined();
});

for (const sample of native.cases)
	it(`flows native title ${sample.referenceName} without changing rich XML`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const book = createWorkbook();
		const verify = (target: typeof chart) => {
			const view = chartView(book, 0, target, () => []);
			const title = chartRichTitleSvg(view, (sample.chartGeometry.widthPt * 4) / 3, {
				measureText: (text, font) => {
					const width = fonts[font.split(',')[0]!]?.widths[text];
					if (width === undefined) throw new Error(`Missing captured width: ${font}, ${text}`);
					return width;
				},
				measureFont: (font) => fonts[font.split(',')[0]!],
			});
			expect(title).toBeDefined();
			const root = parseXml(`<svg>${title!.markup}</svg>`);
			const spans = [...root.getElementsByTagName('tspan')];
			const lines = new Map<string, string>();
			for (const span of spans) {
				const y = span.getAttribute('y')!;
				lines.set(y, (lines.get(y) ?? '') + span.textContent);
			}
			const widths = new Map<string, number>();
			for (const span of spans) {
				const font = `normal ${span.getAttribute('font-weight') ?? 'normal'} ${span.getAttribute('font-size')}px Arial`;
				const y = span.getAttribute('y')!;
				widths.set(y, (widths.get(y) ?? 0) + fonts[font]!.widths[span.textContent ?? '']!);
			}
			expect(
				Math.abs(Math.max(...widths.values()) + 8 - (sample.titleGeometry.widthPt * 4) / 3),
			).toBeLessThan(2);
			expect(Math.abs(title!.height - (sample.titleGeometry.heightPt * 4) / 3)).toBeLessThan(4);
			const expectedLines =
				sample.referenceName === 'wrap-short-480'
					? 1
					: sample.referenceName.includes('long-200')
						? 5
						: sample.referenceName.includes('long-320') ||
							  sample.referenceName.includes('mixed') ||
							  sample.referenceName.includes('word-160') ||
							  sample.referenceName.includes('word-200')
							? 3
							: 2;
			expect(lines.size).toBe(expectedLines);
			expect([...lines.values()].join('').replaceAll(' ', '')).toBe(
				view.title!.replaceAll(' ', ''),
			);
			if (sample.referenceName === 'wrap-word-200')
				expect([...lines.values()]).toEqual(['InternationalBusin', 'essRevenueForec', 'ast']);
			if (sample.referenceName === 'wrap-mixed-320')
				expect([...lines.values()]).toEqual([
					'Revenue growth and',
					'forecast for the entire',
					'international business region',
				]);
			if (sample.referenceName === 'wrap-spaced-480')
				expect(Number(spans[0]!.getAttribute('y'))).toBe(40);
		};
		verify(chart);
		const regenerated = chartXml({ ...chart, chartType: 'line' });
		expect(richXml(regenerated)).toBe(richXml(source));
		verify(parseChart(regenerated, chart.anchor, ''));
	});
