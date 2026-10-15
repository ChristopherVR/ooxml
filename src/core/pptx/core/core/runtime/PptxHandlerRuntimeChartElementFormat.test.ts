import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
	resetChartElementFormat,
	withChartElementFormat,
} from '../../builders/sdk/chart-element-format-operations';
import {
	withChartAxisOptions,
	withChartLegendOptions,
	withChartSeriesOptions,
} from '../../builders/sdk/chart-format-options-operations';
import { PresentationBuilder } from '../../builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../PptxHandler';
import type { PptxChartData } from '../../types';
import type { ChartPptxElement } from '../../types/elements';

/**
 * Load -> chart Format pane edits -> save -> reload. Every edit the panes make
 * (fill with transparency, line width/dash/compound/cap/join/arrows, shadow,
 * glow, soft edges, Reset to Match Style, Series/Axis/Legend options) must
 * land in the right element's XML and come back on reload.
 */

const BAR_CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
	xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<c:chart>
		<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
		<c:autoTitleDeleted val="0"/>
		<c:plotArea>
			<c:layout/>
			<c:barChart>
				<c:barDir val="col"/>
				<c:grouping val="clustered"/>
				<c:varyColors val="0"/>
				<c:ser>
					<c:idx val="0"/>
					<c:order val="0"/>
					<c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>East</c:v></c:pt></c:strCache></c:strRef></c:tx>
					<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>
					<c:invertIfNegative val="0"/>
					<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>
					<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
					<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>
				</c:ser>
				<c:gapWidth val="150"/>
				<c:axId val="111"/>
				<c:axId val="222"/>
			</c:barChart>
			<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/></c:catAx>
			<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/></c:valAx>
			<c:spPr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:effectLst><a:outerShdw blurRad="1"/></a:effectLst></c:spPr>
		</c:plotArea>
		<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
		<c:plotVisOnly val="1"/>
	</c:chart>
	<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>
</c:chartSpace>`;

const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
	<p:cSld><p:spTree>
		<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
		<p:grpSpPr/>
		<p:graphicFrame>
			<p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
			<p:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="3200400"/></p:xfrm>
			<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
				<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/>
			</a:graphicData></a:graphic>
		</p:graphicFrame>
	</p:spTree></p:cSld>
</p:sld>`;

const SLIDE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
	<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

async function buildDeck(): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	zip.file('ppt/slides/slide1.xml', SLIDE_XML);
	zip.file('ppt/slides/_rels/slide1.xml.rels', SLIDE_RELS_XML);
	zip.file('ppt/charts/chart1.xml', BAR_CHART_XML);
	const out = await zip.generateAsync({ type: 'uint8array' });
	return out.buffer as ArrayBuffer;
}

async function loadChart(buffer: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(buffer);
	const chart = data.slides[0].elements.find((e) => e.type === 'chart') as ChartPptxElement;
	expect(chart.chartData).toBeDefined();
	return { handler, data, chart };
}

async function saveWith(edit: (data: PptxChartData) => PptxChartData) {
	const { handler, data, chart } = await loadChart(await buildDeck());
	chart.chartData = edit(chart.chartData!);
	data.slides[0].isDirty = true;
	const bytes = await handler.save(data.slides);
	const zip = await JSZip.loadAsync(bytes);
	// The builder writes empty elements as `<x></x>`; compare in self-closing form.
	const xml = (await zip.file('ppt/charts/chart1.xml')!.async('string')).replace(
		/<([\w:]+)([^<>]*)><\/\1>/gu,
		'<$1$2/>',
	);
	const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
	return { xml, reloaded: await loadChart(buffer as ArrayBuffer) };
}

/** The inner XML of the first `<tag ...>...</tag>` (or self-closing) in `xml`. */
function section(xml: string, tag: string): string {
	const match = new RegExp(`<${tag}[\\s>][\\s\\S]*?</${tag}>`, 'u').exec(xml);
	return match?.[0] ?? '';
}

describe('chart Format pane edits save and reload', () => {
	it('writes legend fill, border and shadow into c:legend/c:spPr in schema order', async () => {
		const { xml } = await saveWith((data) =>
			withChartElementFormat(
				data,
				{ kind: 'legend' },
				{
					fill: { kind: 'solid', color: '#FFF2CC', transparency: 25 },
					line: { kind: 'solid', color: '#BF9000', width: 1.5, dash: 'dash' },
					effects: {
						shadow: { color: '#000000', transparency: 60, blur: 4, distance: 3, angle: 45 },
					},
				},
			),
		);
		const legend = section(xml, 'c:legend');
		expect(legend).toMatch(
			/<c:overlay val="0"\/>\s*<c:spPr><a:solidFill><a:srgbClr val="FFF2CC"><a:alpha val="75000"\/><\/a:srgbClr><\/a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="BF9000"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"/u,
		);
	});

	it('writes the chart area, series, axis line, gridlines and title border', async () => {
		const { xml, reloaded } = await saveWith((data) => {
			let next = withChartElementFormat(
				data,
				{ kind: 'chartArea' },
				{ fill: { kind: 'solid', color: '#DDEBF7', transparency: 10 } },
			);
			next = withChartElementFormat(
				next,
				{ kind: 'series', seriesIndex: 0 },
				{
					fill: { kind: 'solid', color: '#C00000' },
					effects: { glow: { color: '#FFC000', radius: 5, transparency: 40 } },
				},
			);
			next = withChartElementFormat(
				next,
				{ kind: 'axis', axisIndex: 0 },
				{ line: { kind: 'solid', color: '#7F7F7F', width: 2, endArrow: 'triangle' } },
			);
			next = withChartElementFormat(
				next,
				{ kind: 'majorGridlines', axisIndex: 1 },
				{ line: { kind: 'solid', color: '#D9D9D9', compound: 'dbl', cap: 'rnd', join: 'round' } },
			);
			return withChartElementFormat(
				next,
				{ kind: 'title' },
				{ line: { kind: 'solid', color: '#000000', width: 0.75 } },
			);
		});
		// Chart area c:spPr sits after c:chart and before c:txPr.
		expect(xml).toMatch(
			/<\/c:chart>\s*<c:spPr><a:solidFill><a:srgbClr val="DDEBF7"><a:alpha val="90000"\/>/u,
		);
		const ser = section(xml, 'c:ser');
		expect(ser).toMatch(/<\/c:tx>\s*<c:spPr><a:solidFill><a:srgbClr val="C00000"\/>/u);
		expect(ser).toMatch(/<a:effectLst><a:glow rad="63500"><a:srgbClr val="FFC000">/u);
		const catAx = section(xml, 'c:catAx');
		expect(catAx).toMatch(
			/<c:tickLblPos val="nextTo"\/>\s*<c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="7F7F7F"\/><\/a:solidFill><a:tailEnd type="triangle"\/><\/a:ln><\/c:spPr>\s*<c:crossAx/u,
		);
		expect(section(xml, 'c:valAx')).toMatch(
			/<c:majorGridlines><c:spPr><a:ln cmpd="dbl" cap="rnd"><a:solidFill><a:srgbClr val="D9D9D9"\/><\/a:solidFill><a:round\/>/u,
		);
		expect(section(xml, 'c:title')).toMatch(/<c:spPr>.*<a:ln w="9525">/u);

		const chartData = reloaded.chart.chartData!;
		expect(chartData.style?.chartAreaFill?.toUpperCase()).toContain('DDEBF7');
		expect(chartData.series[0].color?.toUpperCase()).toContain('C00000');
		expect(chartData.axes?.[0].spPr?.strokeColor?.toUpperCase()).toContain('7F7F7F');
		expect(chartData.axes?.[1].majorGridlinesSpPr?.strokeColor?.toUpperCase()).toContain('D9D9D9');
	});

	it('Reset to Match Style drops the plot area formatting, effects included', async () => {
		const { xml, reloaded } = await saveWith((data) =>
			resetChartElementFormat(data, { kind: 'plotArea' }),
		);
		const plotArea = section(xml, 'c:plotArea');
		expect(plotArea).not.toContain('F2F2F2');
		expect(plotArea).not.toContain('outerShdw');
		expect(reloaded.chart.chartData!.style?.plotAreaFill).toBeUndefined();
	});

	it('persists Series, Axis and Legend options', async () => {
		const { xml, reloaded } = await saveWith((data) => {
			let next = withChartSeriesOptions(data, { gapWidth: 60, overlap: 20 });
			next = withChartAxisOptions(next, 1, {
				min: 0,
				max: 40,
				majorUnit: 10,
				numberFormat: '0.0',
				majorTickMark: 'cross',
				tickLabelPosition: 'low',
			});
			return withChartLegendOptions(next, { position: 'b', overlay: true });
		});
		expect(xml).toContain('<c:gapWidth val="60"/>');
		expect(xml).toContain('<c:overlap val="20"/>');
		expect(section(xml, 'c:legend')).toMatch(/<c:legendPos val="b"\/>\s*<c:overlay val="1"\/>/u);
		const data = reloaded.chart.chartData!;
		const valAx = data.axes?.find((a) => a.axisType === 'valAx');
		expect(valAx).toMatchObject({ min: 0, max: 40, majorUnit: 10, majorTickMark: 'cross' });
		expect(valAx?.tickLblPos).toBe('low');
		expect(valAx?.numFmt?.formatCode).toBe('0.0');
		expect(data.style?.legendPosition).toBe('b');
		expect(data.style?.legendOverlay).toBe(true);
		expect(data.barGapWidth).toBe(60);
	});

	it('leaves an untouched chart part byte-for-byte', async () => {
		const { handler, data } = await loadChart(await buildDeck());
		data.slides[0].isDirty = true;
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toBe(BAR_CHART_XML);
	});
});
