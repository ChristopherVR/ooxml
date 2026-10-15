import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { chartSummaryView } from './chart-view';
import { chartThemePalette, renderChartSpaceSvg } from './render-space';
import { chartSummaryFromSpace } from './summary-from-space';

const C =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const numbers = (local: string, values: number[]) =>
	`<c:${local}><c:numLit><c:ptCount val="${values.length}"/>${values
		.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
		.join('')}</c:numLit></c:${local}>`;

const axes =
	'<c:axId val="1"/><c:axId val="2"/></GROUP><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>' +
	'<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>';

const TRENDLINE =
	'<c:trendline><c:name>Fit</c:name><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr><c:trendlineType val="linear"/></c:trendline>';
const ERROR_BARS =
	'<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>';

const part = (group: string, head: string, analysis: string, grouping = 'clustered') =>
	`<c:chartSpace ${C}><c:chart><c:plotArea><c:${group}>${head}<c:grouping val="${grouping}"/>` +
	`<c:ser><c:idx val="0"/><c:order val="0"/>${analysis}${numbers('val', [2, 4, 6, 8])}</c:ser>` +
	axes.replace('GROUP', `c:${group}`) +
	'</c:plotArea></c:chart></c:chartSpace>';

const svgOf = (xml: string) =>
	renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 400, height: 300 });

describe('trendlines and error bars in the shared chart painter', () => {
	it('carries them from the chart part to the summary and the view', () => {
		const space = parseChartSpace(
			part('lineChart', '', TRENDLINE.replace('<c:name>Fit</c:name>', '') + ERROR_BARS, 'standard'),
		).chartSpace;
		const summary = chartSummaryFromSpace(space);
		expect(summary.series[0]?.trendlines).toEqual([
			{ kind: 'linear', line: expect.objectContaining({ widthEmu: 25400, dash: 'dash' }) },
		]);
		expect(summary.series[0]?.errorBars).toEqual([
			{ direction: 'y', barType: 'both', valueType: 'fixedVal', value: 1 },
		]);
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		const series = view.series[0];
		expect(series?.trendlines?.[0]?.stroke).toEqual({
			color: '#FF0000',
			width: 2 * (4 / 3),
			dash: 'dash',
		});
		expect(series?.errorBars?.[0]).toMatchObject({ endCap: true, stroke: { color: '#595959' } });
	});

	it('draws a fitted line and capped error bars over a line chart', () => {
		const svg = svgOf(part('lineChart', '', TRENDLINE + ERROR_BARS, 'standard'));
		const trend = /<polyline points="([^"]+)"[^>]*data-chart-part="trendline"/.exec(svg);
		expect(trend).not.toBeNull();
		expect(trend?.[0]).toContain('stroke="#FF0000"');
		expect(trend?.[0]).toMatch(/stroke-dasharray="10\.67 8"/);
		// A perfect linear series: the fitted samples are collinear.
		const points = trend![1]!.split(' ').map((pair) => pair.split(',').map(Number));
		const [[x0, y0], [x1, y1]] = [points[0]!, points.at(-1)!];
		for (const [x, y] of points)
			expect(Math.abs((y! - y0!) * (x1! - x0!) - (y1! - y0!) * (x! - x0!))).toBeLessThan(
				(x1! - x0!) * 0.02,
			);
		// Four points, two sides each, each with a cap.
		expect(svg.match(/<line [^>]*stroke="#595959" stroke-width="1"\/>/g)).toHaveLength(16);
	});

	it('defaults a trendline to the series colour, dotted, and skips stacked charts', () => {
		const plain = svgOf(
			part(
				'barChart',
				'<c:barDir val="col"/>',
				'<c:trendline><c:trendlineType val="movingAvg"/></c:trendline>',
			),
		);
		expect(plain).toMatch(
			/<polyline [^>]*stroke="#4472C4" stroke-width="2" stroke-dasharray="2 2"/,
		);
		const stacked = svgOf(
			part('barChart', '<c:barDir val="col"/>', TRENDLINE + ERROR_BARS, 'stacked'),
		);
		expect(stacked).not.toContain('data-chart-part="trendline"');
	});

	it('fits a scatter trendline against the x values and draws x error bars', () => {
		const xml =
			`<c:chartSpace ${C}><c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="marker"/>` +
			`<c:ser><c:idx val="0"/><c:order val="0"/>${TRENDLINE}` +
			'<c:errBars><c:errDir val="x"/><c:errBarType val="plus"/><c:errValType val="percentage"/><c:val val="10"/><c:noEndCap val="1"/></c:errBars>' +
			`${numbers('xVal', [1, 2, 4, 8])}${numbers('yVal', [3, 5, 9, 17])}</c:ser>` +
			'<c:axId val="1"/><c:axId val="2"/></c:scatterChart><c:valAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>' +
			'<c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>';
		const svg = svgOf(xml);
		expect(svg).toContain('data-chart-part="trendline"');
		// One horizontal bar per point, no caps.
		const bars = [
			...svg.matchAll(
				/<line x1="([\d.]+)" y1="([\d.]+)" x2="([\d.]+)" y2="([\d.]+)" stroke="#595959"/g,
			),
		];
		expect(bars).toHaveLength(4);
		for (const [, x1, y1, x2, y2] of bars) {
			expect(y1).toBe(y2);
			expect(Number(x2)).toBeGreaterThan(Number(x1));
		}
	});
});
