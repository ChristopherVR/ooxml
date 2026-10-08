import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../builders/sdk/PresentationBuilder';
import { PptxHandler } from '../PptxHandler';
import type { ChartPptxElement } from '../types/elements';
import { readChartPartModel, repairChartPartPrefixes } from './chart-part-repair';

const C15 = 'http://schemas.microsoft.com/office/drawing/2012/chart';
const C16 = 'http://schemas.microsoft.com/office/drawing/2014/chart';

/**
 * A chart part as this library's chart writers saved it before the c15/c16 fix: the filtered-series
 * writer put `xmlns:c15`/`xmlns:c16` in as ELEMENTS and wrote the hidden series' name as `c15:v`;
 * the data-label-range writer did the same for its own extension. The root declares only `c`, `a`
 * and `r`, so `c15` and `c16` are used undeclared.
 */
const OLD_WRITER_PART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Series C</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"><xmlns:c15>${C15}</xmlns:c15><c15:showDataLabelsRange val="1"/><c15:datalabelsRange><c15:f>Sheet1!$D$2:$D$4</c15:f><c15:dlblRangeCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Low</c:v></c:pt><c:pt idx="1"><c:v>Mid</c:v></c:pt><c:pt idx="2"><c:v>High</c:v></c:pt></c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst></c:dLbls>
<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>4</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>6</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
<c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/>
<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}"><xmlns:c15>${C15}</xmlns:c15><c15:filteredBarSeries><c15:ser><c:idx val="1"/><c:order val="1"/><c:tx><c15:v>Series A</c15:v></c:tx><c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val><c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}"><xmlns:c16>${C16}</xmlns:c16><c16:uniqueId val="{00000001-0000-0000-0000-000000000000}"/></c:ext></c:extLst></c15:ser></c15:filteredBarSeries></c:ext></c:extLst></c:barChart>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
</c:plotArea></c:chart></c:chartSpace>`;

describe('repairChartPartPrefixes', () => {
	it('removes xmlns elements, maps c15:v to c:v and declares the used Office prefixes', () => {
		const repair = repairChartPartPrefixes(OLD_WRITER_PART)!;
		expect(repair.declared.sort()).toEqual(['c15', 'c16']);
		expect(repair.rewroteOwnOutput).toBe(true);
		expect(repair.text).not.toContain('<xmlns:');
		expect(repair.text).not.toContain('c15:v');
		expect(repair.text).toContain(`xmlns:c15="${C15}"`);
	});

	it('has nothing to repair in a well-formed part', () => {
		const repaired = repairChartPartPrefixes(OLD_WRITER_PART)!.text;
		expect(repairChartPartPrefixes(repaired)).toBeUndefined();
	});
});

describe('readChartPartModel', () => {
	it('loads a part the old writers saved, with the series intact, and reports the repair', () => {
		const { chartSpace, repairIssue } = readChartPartModel(OLD_WRITER_PART);
		expect(repairIssue?.code).toBe('CHART_PREFIX_REPAIRED');
		const series = chartSpace?.plotArea.groups[0]?.series ?? [];
		expect(series.map((entry) => entry.tx?.text)).toEqual(['Series C']);
		expect(series[0]?.values?.cache?.points.map((point) => point.value)).toEqual(['4', '5', '6']);
		expect(chartSpace?.plotArea.groups[0]?.extLst).toContain('Series A');
	});

	it('does not report anything for a well-formed part', () => {
		const repaired = repairChartPartPrefixes(OLD_WRITER_PART)!.text;
		expect(readChartPartModel(repaired).repairIssue).toBeUndefined();
	});

	it('still rejects a part malformed in an unrelated way', () => {
		// An unknown undeclared prefix is not one the repair declares.
		expect(
			readChartPartModel(OLD_WRITER_PART.replace('<c:gapWidth', '<foo:x/><c:gapWidth')),
		).toEqual({});
		// Not well-formed at all.
		expect(readChartPartModel(OLD_WRITER_PART.replace('</c:plotArea>', ''))).toEqual({});
	});
});

const SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4572000" cy="3200400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></p:graphicFrame>
</p:spTree></p:cSld></p:sld>`;
const SLIDE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`;

describe('loading a deck whose chart part the old writers saved', () => {
	it('loads the chart with its series, hidden series and label range intact', async () => {
		const { handler, data, createSlide } = await PresentationBuilder.create();
		data.slides.push(createSlide('Blank').build());
		const zip = await JSZip.loadAsync(await handler.save(data.slides));
		zip.file('ppt/slides/slide1.xml', SLIDE);
		zip.file('ppt/slides/_rels/slide1.xml.rels', SLIDE_RELS);
		zip.file('ppt/charts/chart1.xml', OLD_WRITER_PART);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await new PptxHandler().load(bytes.slice().buffer as ArrayBuffer);
		const chart = loaded.slides[0]?.elements.find(
			(element): element is ChartPptxElement => element.type === 'chart',
		);
		const chartData = chart?.chartData;
		expect(chartData?.chartType).toBe('bar');
		expect(chartData?.categories).toEqual(['Q1', 'Q2', 'Q3']);
		expect(chartData?.series.map((entry) => [entry.name, entry.values])).toEqual([
			['Series C', [4, 5, 6]],
		]);
		expect(chartData?.filteredSeries?.map((entry) => [entry.name, entry.values])).toEqual([
			['Series A', [1, 2, 3]],
		]);
		expect(loaded.warnings?.map((warning) => warning.code)).toContain('CHART_PREFIX_REPAIRED');
		expect(chartData?.series[0]?.dataLabelOptions?.dataLabelsRange).toEqual({
			formula: 'Sheet1!$D$2:$D$4',
			cache: ['Low', 'Mid', 'High'],
		});
	});
});
