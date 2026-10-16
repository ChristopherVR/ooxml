/**
 * A one-chart deck for chart Format pane round-trip tests: build, load, edit,
 * save and reload, plus a helper that cuts one element out of the saved XML.
 */
import JSZip from 'jszip';
import { expect } from 'vitest';

import { PresentationBuilder } from '../../builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../PptxHandler';
import type { PptxChartData } from '../../types';
import type { ChartPptxElement } from '../../types/elements';

export const BAR_CHART_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

export const SLIDE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

export const SLIDE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
	<Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;

export async function buildDeck(chartXml = BAR_CHART_XML): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	zip.file('ppt/slides/slide1.xml', SLIDE_XML);
	zip.file('ppt/slides/_rels/slide1.xml.rels', SLIDE_RELS_XML);
	zip.file('ppt/charts/chart1.xml', chartXml);
	const out = await zip.generateAsync({ type: 'uint8array' });
	return out.buffer as ArrayBuffer;
}

export async function loadChart(buffer: ArrayBuffer) {
	const handler = new PptxHandler();
	const data = await handler.load(buffer);
	const chart = data.slides[0].elements.find((e) => e.type === 'chart') as ChartPptxElement;
	expect(chart.chartData).toBeDefined();
	return { handler, data, chart };
}

export async function saveWith(
	edit: (data: PptxChartData) => PptxChartData,
	chartXml = BAR_CHART_XML,
) {
	const { handler, data, chart } = await loadChart(await buildDeck(chartXml));
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
	return { xml, zip, reloaded: await loadChart(buffer as ArrayBuffer) };
}

/** The inner XML of the first `<tag ...>...</tag>` (or self-closing) in `xml`. */
export function section(xml: string, tag: string): string {
	const match = new RegExp(`<${tag}[\\s>][\\s\\S]*?</${tag}>`, 'u').exec(xml);
	return match?.[0] ?? '';
}
