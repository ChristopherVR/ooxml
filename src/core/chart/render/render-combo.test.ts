import { describe, expect, it } from 'vitest';
import { parseChartSpace } from '../parse-space';
import { chartSummaryFromSpace } from './summary-from-space';
import { chartSpaceKind, chartThemePalette, renderChartSpaceSvg } from './render-space';
import { chartSummaryView } from './chart-view';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';

const series = (idx: number, name: string, values: number[]) =>
	`<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:strRef><c:f>S!$${'BCD'[idx]}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
	`<c:cat><c:strRef><c:f>S!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/>${['Q1', 'Q2', 'Q3', 'Q4']
		.map((c, i) => `<c:pt idx="${i}"><c:v>${c}</c:v></c:pt>`)
		.join('')}</c:strCache></c:strRef></c:cat>` +
	`<c:val><c:numRef><c:f>S!$B$2:$B$5</c:f><c:numCache><c:ptCount val="4"/>${values
		.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
		.join('')}</c:numCache></c:numRef></c:val></c:ser>`;

/** A clustered column group on the primary axes plus a line group on a secondary value axis. */
const comboXml = (secondary = true) => `<c:chartSpace ${C}><c:chart><c:plotArea><c:layout/>
<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${series(0, 'Revenue', [100, 140, 120, 180])}<c:gapWidth val="100"/><c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:lineChart><c:grouping val="standard"/>${series(1, 'Margin', [0.1, 0.2, 0.15, 0.3])}<c:marker val="1"/><c:axId val="${secondary ? 3 : 1}"/><c:axId val="${secondary ? 4 : 2}"/></c:lineChart>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
${
	secondary
		? `<c:catAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:crossAx val="4"/></c:catAx>
<c:valAx><c:axId val="4"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:crossAx val="3"/><c:crosses val="max"/></c:valAx>`
		: ''
}
</c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>`;

describe('combination charts', () => {
	it('keeps the groups after the first, flagging the one on another value axis', () => {
		const { chartSpace } = parseChartSpace(comboXml());
		const summary = chartSummaryFromSpace(chartSpace);
		expect(summary.chartType).toBe('column');
		expect(summary.series.map((s) => s.name)).toEqual(['Revenue']);
		expect(summary.extraGroups).toHaveLength(1);
		expect(summary.extraGroups?.[0]).toMatchObject({ chartType: 'line', secondary: true });
		expect(summary.extraGroups?.[0]?.series[0]?.name).toBe('Margin');
		const shared = chartSummaryFromSpace(parseChartSpace(comboXml(false)).chartSpace);
		expect(shared.extraGroups?.[0]?.secondary).toBe(false);
	});

	it('scales the secondary axis from its own groups and the primary one from the rest', () => {
		const summary = chartSummaryFromSpace(parseChartSpace(comboXml()).chartSpace);
		const view = chartSummaryView(summary, chartThemePalette(undefined));
		expect(view.valueAxis?.max).toBeGreaterThanOrEqual(180);
		expect(view.secondaryAxis?.max).toBeLessThanOrEqual(0.5);
		expect(view.extraGroups?.[0]?.series[0]?.color).toBe('#ED7D31');
		const onPrimary = chartSummaryView(
			chartSummaryFromSpace(parseChartSpace(comboXml(false)).chartSpace),
			chartThemePalette(undefined),
		);
		expect(onPrimary.secondaryAxis).toBeUndefined();
		expect(onPrimary.valueAxis?.max).toBeGreaterThanOrEqual(180);
	});

	it('draws bars, the line, a right-hand axis and both legend entries', () => {
		const { chartSpace } = parseChartSpace(comboXml());
		expect(chartSpaceKind(chartSpace)).toEqual({ type: 'column', drawn: true });
		const svg = renderChartSpaceSvg(chartSpace, { width: 480, height: 300 });
		expect(svg.match(/<g data-chart-series="0"/g)).toHaveLength(4);
		expect(svg).toContain('<polyline');
		// Series index 1 uses the native automatic square marker.
		expect(svg.match(/width="5" height="5"/g)).toHaveLength(4);
		// Secondary tick labels: percent-scale fractions sit right of the plot.
		expect(svg).toContain('>0.3<');
		expect(svg).toContain('>Revenue<');
		expect(svg).toContain('>Margin<');
		// The line is in the second accent colour.
		expect(svg).toContain('stroke="#ED7D31"');
	});

	it('does not layer groups over a horizontal bar first group', () => {
		const xml = comboXml().replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>');
		const summary = chartSummaryFromSpace(parseChartSpace(xml).chartSpace);
		expect(summary.chartType).toBe('bar');
		expect(summary.extraGroups).toBeUndefined();
	});
});
