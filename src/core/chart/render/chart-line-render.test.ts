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
