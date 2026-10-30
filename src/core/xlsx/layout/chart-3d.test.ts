import { expect, it } from 'vitest';
import { renderChartSvg } from '../../chart/render/chart-svg';
import { parseChart } from '../read/chart';
import { createWorkbook } from '../workbook';
import { chartView } from './chart-view';

const xml = (
	group: string,
) => `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D><c:plotArea><c:layout/>${group}
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;

const series = `<c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>6</c:v></c:pt><c:pt idx="2"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser>`;

function svgOf(group: string): string {
	const chart = parseChart(
		xml(group),
		{ from: { row: 0, col: 0, rowOffset: 0, colOffset: 0 } },
		'xl/charts/chart1.xml',
	);
	return renderChartSvg(
		chartView(createWorkbook(), 0, chart, () => []),
		400,
		300,
	);
}

it('draws a 3-D column chart of a workbook in 3-D', () => {
	const svg = svgOf(
		`<c:bar3DChart><c:barDir val="col"/><c:grouping val="clustered"/>${series}<c:axId val="1"/><c:axId val="2"/></c:bar3DChart>`,
	);
	expect(svg).toContain('<polygon');
	expect(svg.match(/<g data-chart-series/g)).toHaveLength(3);
});

it('draws 3-D line, area and pie charts of a workbook in 3-D', () => {
	for (const group of [
		`<c:line3DChart><c:grouping val="standard"/>${series}<c:axId val="1"/><c:axId val="2"/></c:line3DChart>`,
		`<c:area3DChart><c:grouping val="standard"/>${series}<c:axId val="1"/><c:axId val="2"/></c:area3DChart>`,
		`<c:pie3DChart><c:varyColors val="1"/>${series}</c:pie3DChart>`,
	])
		expect(svgOf(group)).toContain('<polygon');
});
