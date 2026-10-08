import { expect, it } from 'vitest';
import native from './excel-chart-title-spacing.json';
import { parseChart } from '../xlsx/read/chart';
import { createWorkbook } from '../xlsx/workbook';
import { chartView } from '../xlsx/layout/chart-view';
import { chartRichTitleSvg } from './render/chart-svg-title-text';
import { chartXml } from '../xlsx/write/chart';
import { NS, first, parseXml, buildXml } from '../xml';

const richXml = (xml: string) =>
	buildXml(
		first(
			first(first(first(parseXml(xml).documentElement, 'chart', NS.c), 'title', NS.c), 'tx', NS.c),
			'rich',
			NS.c,
		)!,
	);

for (const sample of native.cases)
	it(`renders and preserves native title spacing ${sample.referenceName}`, () => {
		const source = sample.parts['xl/charts/chart1.xml'];
		const chart = parseChart(source, { from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } }, '');
		const book = createWorkbook();
		const verify = (target: typeof chart) => {
			const view = chartView(book, 0, target, () => []);
			expect(view.titleText).toBeDefined();
			const title = chartRichTitleSvg(view, 640, {
				measureFont: () => ({ ascent: 22, descent: 5 }),
			})!;
			expect(Math.abs(title.height - (sample.titleGeometry.heightPt * 4) / 3)).toBeLessThan(2);
			const spacing = view.titleText!.paragraphs[0]!;
			if (sample.referenceName.includes('percent-150'))
				expect(spacing.lineSpacing).toEqual({ unit: 'percent', value: 1.5 });
			if (sample.referenceName === 'spacing-points-30')
				expect(spacing.lineSpacing).toEqual({ unit: 'points', value: 30 });
			if (sample.referenceName.includes('before-after')) {
				expect(spacing.spaceBefore).toEqual({ unit: 'points', value: 6 });
				expect(spacing.spaceAfter).toEqual({ unit: 'points', value: 8 });
			}
			if (sample.referenceName === 'spacing-paragraphs')
				expect(view.titleText!.paragraphs).toHaveLength(2);
			const spans = [...parseXml(`<svg>${title.markup}</svg>`).getElementsByTagName('tspan')];
			const firstY = Number(spans[0]!.getAttribute('y'));
			const before =
				sample.referenceName === 'spacing-before-percent-50'
					? 13.5
					: [
								'spacing-before-6',
								'spacing-before-after',
								'spacing-paragraphs',
								'spacing-single-before-after',
						  ].includes(sample.referenceName)
						? 8
						: 0;
			const after =
				sample.referenceName === 'spacing-after-percent-25'
					? 6.75
					: [
								'spacing-after-8',
								'spacing-before-after',
								'spacing-paragraphs',
								'spacing-single-before-after',
						  ].includes(sample.referenceName)
						? 32 / 3
						: 0;
			const pitch = sample.referenceName.includes('percent-150')
				? 40.5
				: sample.referenceName === 'spacing-percent-75'
					? 20.25
					: sample.referenceName === 'spacing-points-30'
						? 40
						: 27;
			expect(firstY).toBeCloseTo(32 + before + (pitch - 27) / 2, 2);
			if (!sample.referenceName.includes('single')) {
				const lastY = Number(spans.at(-1)!.getAttribute('y'));
				expect(lastY - firstY).toBeCloseTo(pitch + before + after, 2);
			}
		};
		verify(chart);
		const regenerated = chartXml({ ...chart, chartType: 'line' });
		expect(richXml(regenerated)).toBe(richXml(source));
		verify(parseChart(regenerated, chart.anchor, ''));
	});
