import { chartSummaryView } from './chart-view';
import { chartSummaryFromSpace } from './summary-from-space';
import { chartThemePalette } from './render-space';
import { renderChartSvg } from './chart-svg';
import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { renderChartSpaceSvg } from './render-space';
import { smoothLinePath, sampleSmoothLine } from '../line-path';
const ns =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const points = (values: string[]) =>
	values.map((v, i) => (v === '' ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('');
const fixture = (group: string, options = '', values = ['1', '3', '2'], style = '') =>
	`<c:chartSpace ${ns}><c:chart><c:plotArea><c:${group}>${style}<c:ser><c:idx val="0"/><c:order val="0"/>${options}<c:cat><c:numLit>${points(['1', '2', '3', '4', '5'])}</c:numLit></c:cat><c:val><c:numLit>${points(values)}</c:numLit></c:val></c:ser></c:${group}></c:plotArea></c:chart></c:chartSpace>`;
const render = (xml: string) =>
	renderChartSpaceSvg(parseChartSpace(xml).chartSpace, { width: 400, height: 300 });
describe('shared series paths', () => {
	it('uses the same curve as PPTX, including short runs and exact endpoints', () => {
		const points = [
			{ x: 0, y: 2 },
			{ x: 4, y: 8 },
			{ x: 10, y: 1 },
		];
		expect(smoothLinePath(points)).toContain(' C');
		expect(sampleSmoothLine(points).at(-1)).toEqual(points[2]);
		expect(smoothLinePath(points.slice(0, 2))).toEqual('M0.00,2.00 L4.00,8.00');
		expect(smoothLinePath([])).toEqual('');
	});
	it('inherits group smoothing and respects an explicit series override', () => {
		expect(render(fixture('lineChart', '', undefined, '<c:smooth val="1"/>'))).toMatch(
			/<path d="M[^" ]+ C/u,
		);
		expect(
			render(fixture('lineChart', '<c:smooth val="0"/>', undefined, '<c:smooth val="1"/>')),
		).toContain('<polyline');
	});
	it('keeps null gaps as separate runs and applies authored stroke width and dash', () => {
		const svg = render(
			fixture(
				'lineChart',
				'<c:smooth val="1"/><c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>',
				['1', '2', '', '3', '4'],
			),
		);
		expect(svg.match(/<path d=/gu) ?? []).toHaveLength(2);
		expect(svg).not.toMatch(/d="[^"]* C/u);
		expect(svg).toContain('stroke="#FF0000" stroke-width="3" stroke-dasharray="12 9"');
	});
	it('draws scatter lines only for line styles, including smooth and hidden outlines', () => {
		expect(
			render(fixture('scatterChart', '', undefined, '<c:scatterStyle val="smoothMarker"/>')),
		).toContain('<path');
		expect(
			render(fixture('scatterChart', '', undefined, '<c:scatterStyle val="marker"/>')),
		).not.toContain('<polyline');
		expect(render(fixture('lineChart', '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'))).toContain(
			'stroke="none" stroke-width="0"',
		);
	});
});

describe('calendar category axis', () => {
	const dateFixture = (orientation = 'minMax', origin1904 = false, min = 0) =>
		fixture('lineChart')
			.replace(
				points(['1', '2', '3', '4', '5']),
				points([String(min), String(min + 1), String(min + 10)]),
			)
			.replace('<c:chart>', `${origin1904 ? '<c:date1904 val="1"/>' : ''}<c:chart>`)
			.replace(
				'</c:plotArea>',
				`<c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="${orientation}"/><c:min val="${min}"/><c:max val="${min + 10}"/></c:scaling><c:axPos val="b"/><c:numFmt formatCode="yyyy-mm-dd" sourceLinked="0"/><c:majorUnit val="5"/><c:majorTimeUnit val="days"/></c:dateAx></c:plotArea>`,
			);
	it('spaces irregular dates continuously and reverses them without changing source data', () => {
		const xy = (svg: string) =>
			(/<polyline points="([^"]+)"/u.exec(svg)?.[1] ?? '')
				.split(' ')
				.map((p) => Number(p.split(',')[0]));
		const normal = xy(render(dateFixture()));
		expect(normal).toHaveLength(3);
		expect((normal[2]! - normal[1]!) / (normal[1]! - normal[0]!)).toBeCloseTo(9, 1);
		const reverse = xy(render(dateFixture('maxMin')));
		expect(reverse[0]).toBeGreaterThan(reverse[1]!);
		expect(reverse[1]).toBeGreaterThan(reverse[2]!);
	});

	it.each([false, true])(
		'paints calendar ticks above an opaque plot for horizontal=%s',
		(horizontal) => {
			const space = parseChartSpace(dateFixture('minMax', true)).chartSpace;
			const axis = space.plotArea.axes[0]!;
			axis.majorTickMark = 'cross';
			if (horizontal) {
				space.plotArea.groups[0]!.kind = 'bar';
				space.plotArea.groups[0]!.barDirection = 'bar';
				axis.position = 'l';
			}
			const view = chartSummaryView(chartSummaryFromSpace(space), chartThemePalette(undefined));
			view.appearance = { ...view.appearance, plotArea: { fillColor: '#123456' } };
			const svg = renderChartSvg(view, 400, 300);
			const background = svg.indexOf('data-chart-part="plotArea"');
			const tick = svg.indexOf('data-chart-part="dateTick"');
			expect(background).toBeGreaterThan(0);
			expect(tick).toBeGreaterThan(background);
		},
	);
	it('suppresses calendar axis labels when table headers replace them', () => {
		const source = dateFixture('minMax', true).replace('</c:plotArea>', '<c:dTable/></c:plotArea>');
		expect(render(source)).not.toContain('1904-01-01');
		expect(render(source)).toContain('data-chart-part="dataTable"');
	});
	it('uses the chart date system and authored format for calendar ticks', () => {
		expect(render(dateFixture('minMax', true))).toContain('1904-01-01');
		expect(render(dateFixture('minMax', false, 45292))).toContain('2024-01-01');
	});
});

describe('chart data tables', () => {
	const tableFixture = (flags = '') =>
		fixture('lineChart')
			.replace('</c:plotArea>', `<c:dTable>${flags}</c:dTable></c:plotArea>`)
			.replace('<c:val><c:numLit>', '<c:val><c:numLit><c:formatCode>0.0%</c:formatCode>');
	const table = (svg: string) =>
		/<g data-chart-part="dataTable">([\s\S]*?)<\/g>/u.exec(svg)?.[1] ?? '';
	it('reserves room and draws category headers, keys and formatted values without duplicate labels', () => {
		const svg = render(tableFixture());
		const content = table(svg);
		expect(content).toContain('100.0%');
		expect(content).toContain('Series1');
		expect(content).toContain('<line');
		expect(svg.match(/text-anchor="middle"[^>]*>1<\/text>/gu)).toHaveLength(1);
		expect(svg).not.toContain('NaN');
	});
	it('honors border and key switches', () => {
		const content = table(
			render(
				tableFixture(
					'<c:showHorzBorder val="0"/><c:showVertBorder val="0"/><c:showOutline val="0"/><c:showKeys val="0"/>',
				),
			),
		);
		expect(content).not.toContain('<line');
		expect(content).not.toContain('Series1');
		expect(content).toContain('100.0%');
	});
	it('keeps a small viewport finite and leaves absent tables automatic', () => {
		const space = parseChartSpace(tableFixture()).chartSpace;
		expect(renderChartSpaceSvg(space, { width: 90, height: 70 })).not.toMatch(/NaN|Infinity/u);
		expect(render(fixture('lineChart'))).not.toContain('data-chart-part="dataTable"');
	});
});
