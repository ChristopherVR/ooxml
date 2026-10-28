import { chartObjectFromSpace } from '../../xlsx/read/chart';
import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { renderChartSpaceSvg } from './render-space';
import { resolveManualPlotRect } from '../manual-layout';

const xml = (layout: string) =>
	`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:layout>${layout}</c:layout><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>`;
const render = (layout: string) =>
	renderChartSpaceSvg(parseChartSpace(xml(layout)).chartSpace, { width: 400, height: 300 });
const manual = (target: string) =>
	`<c:manualLayout><c:layoutTarget val="${target}"/><c:xMode val="edge"/><c:yMode val="edge"/><c:x val="0.25"/><c:y val="0.2"/><c:w val="0.5"/><c:h val="0.5"/></c:manualLayout>`;

describe('manual plot layout', () => {
	it('retains manual placement in the Excel and embedded workbook chart summary', () => {
		const space = parseChartSpace(xml(manual('inner'))).chartSpace;
		const chart = chartObjectFromSpace(space, {
			anchor: { from: { col: 0, row: 0, colOffset: 0, rowOffset: 0 } },
		});
		expect(chart.plotLayout).toEqual(space.plotArea.layout);
	});
	it('places the actual series inside an inner plot rectangle', () => {
		const svg = render(manual('inner'));
		expect(svg).toContain('150,150');
		expect(svg).toContain('250,90');
		expect(svg).not.toEqual(render(''));
	});
	it('insets outer layouts for axis bands while retaining explicit inner bounds', () => {
		const auto = { x: 30, y: 10, width: 300, height: 250 };
		const bands = { left: 20, top: 5, right: 10, bottom: 25 };
		const layout = {
			x: 0.25,
			y: 0.2,
			width: 0.5,
			height: 0.5,
			xMode: 'edge',
			yMode: 'edge',
		} as const;
		expect(
			resolveManualPlotRect(
				{ ...layout, layoutTarget: 'inner' },
				{ width: 400, height: 300 },
				auto,
				bands,
			),
		).toEqual({ x: 100, y: 60, width: 200, height: 150 });
		expect(resolveManualPlotRect(layout, { width: 400, height: 300 }, auto, bands)).toEqual({
			x: 120,
			y: 65,
			width: 170,
			height: 120,
		});
	});
	it('retains omitted automatic fields and resolves factor offsets and edge sizes', () => {
		expect(
			resolveManualPlotRect(
				{ x: 0.1, width: 0.8, widthMode: 'edge', layoutTarget: 'inner' },
				{ width: 400, height: 300 },
				{ x: 30, y: 10, width: 300, height: 250 },
				{ left: 0, top: 0, right: 0, bottom: 0 },
			),
		).toEqual({ x: 70, y: 10, width: 250, height: 250 });
		expect(render('<c:manualLayout/>')).toEqual(render(''));
	});
});
