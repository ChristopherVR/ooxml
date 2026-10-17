import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { chartSummaryFromSpace } from './summary-from-space';
import { chartSummaryView } from './chart-view';
import { chartThemePalette, renderChartSpaceSvg } from './render-space';
import { chartAxisRatio } from './chart-view-axes';

const xml = (valueOptions = '', categoryOptions = '', scaling = '', swapped = false) => {
	const cat = `<c:catAx><c:axId val="10"/><c:scaling/><c:axPos val="b"/><c:crossAx val="20"/>${categoryOptions}</c:catAx>`;
	const val = `<c:valAx><c:axId val="20"/><c:scaling>${scaling}</c:scaling><c:axPos val="l"/><c:crossAx val="10"/>${valueOptions}</c:valAx>`;
	return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>100</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart>${swapped ? val + cat : cat + val}</c:plotArea></c:chart></c:chartSpace>`;
};
const view = (source: string) =>
	chartSummaryView(
		chartSummaryFromSpace(parseChartSpace(source).chartSpace),
		chartThemePalette(undefined),
	);
const render = (source: string) =>
	renderChartSpaceSvg(parseChartSpace(source).chartSpace, { width: 480, height: 300 });

describe('shared native chart axes', () => {
	it('draws only the gridlines declared on the axis, including native line styling', () => {
		const plain = render(xml());
		const styled = render(
			xml(
				'<c:majorGridlines><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:majorGridlines>',
			),
		);
		expect(plain).not.toContain('stroke="#D9D9D9"');
		expect(styled).toContain('stroke="#FF0000"');
		expect(styled).toContain('stroke-width="2"');
		expect(styled).toContain('stroke-dasharray="2 4"');
	});
	it('honours bounds, major units and axis IDs regardless of XML order', () => {
		const model = view(
			xml('<c:majorUnit val="25"/>', '', '<c:min val="0"/><c:max val="100"/>', true),
		);
		expect(model.axes?.value?.source.id).toBe(20);
		expect(model.valueAxis?.ticks).toEqual([0, 25, 50, 75, 100]);
	});
	it('draws category minor gridlines without requiring major gridlines', () => {
		const svg = render(
			xml(
				'',
				'<c:minorGridlines><c:spPr><a:ln><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill></a:ln></c:spPr></c:minorGridlines><c:minorUnit val="0.5"/>',
			),
		);
		expect(svg).toContain('stroke="#0000FF"');
	});
	it('uses the same logarithmic mapping for radar vertices', () => {
		const source = xml(
			'',
			'',
			'<c:logBase val="10"/><c:min val="1"/><c:max val="100"/>',
		).replaceAll('barChart', 'radarChart');
		const svg = render(source);
		const points = /<polygon points="([^"]+)" fill="none" stroke="#4472C4"/
			.exec(svg)?.[1]
			?.split(' ');
		expect(points).toHaveLength(3);
		const coords = points!.map((point) => point.split(',').map(Number));
		const [cx, cy] = coords[0]!;
		const distance = (point: number[]) => Math.hypot(point[0]! - cx!, point[1]! - cy!);
		// 1 is at the center; 10 is halfway to the outer point at 100.
		expect(distance(coords[1]!) / distance(coords[2]!)).toBeCloseTo(0.5, 3);
	});
	it('maps logarithmic powers evenly and reverses the value axis', () => {
		const model = view(
			xml(
				'',
				'',
				'<c:logBase val="10"/><c:min val="1"/><c:max val="100"/><c:orientation val="maxMin"/>',
			),
		);
		expect(model.valueAxis?.ticks).toEqual([1, 10, 100]);
		const scale = model.valueAxis!;
		expect(chartAxisRatio(1, scale, model.axes?.value)).toBe(1);
		expect(chartAxisRatio(10, scale, model.axes?.value)).toBeCloseTo(0.5);
		expect(chartAxisRatio(100, scale, model.axes?.value)).toBe(0);
	});
	it('suppresses deleted axis labels without suppressing its gridlines', () => {
		const svg = render(xml('<c:delete val="1"/><c:majorGridlines/>', '<c:delete val="1"/>'));
		expect(svg).not.toContain('<text ');
		expect(svg).toContain('stroke="#D9D9D9"');
	});
	it('uses the native numeric format rather than a raw fraction', () => {
		const svg = render(
			xml(
				'<c:numFmt formatCode="0%" sourceLinked="0"/><c:majorUnit val="0.5"/>',
				'',
				'<c:min val="0"/><c:max val="1"/>',
			),
		);
		expect(svg).toContain('>50%</text>');
	});
});
