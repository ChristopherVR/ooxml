import { describe, expect, it } from 'vitest';
import { renderChartSummarySvg } from './render-space';
import { chartExSummary } from './summary-from-chartex';
import { parseChartExSpace } from '../ex/parse';
import { writeChartExSpaceOverSource } from '../ex/keep-source';
const cx = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
// Published Bing PCA vector, transformed into a small test polygon; no provider boundaries.
const clear = `<cx:clear xmlns:cx="${cx}"><cx:geoData entityId="sample" entityName="Sample"><cx:geoPolygons><cx:geoPolygon numPoints="4" pcaRings="1,vx1vilihnM6hR7mEl2Q"/></cx:geoPolygons><cx:copyrights><cx:copyright>Test attribution</cx:copyright></cx:copyrights></cx:geoData></cx:clear>`;
const xml = (cache = clear, projection = 'mercator') =>
	`<cx:chartSpace xmlns:cx="${cx}"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="3"><cx:pt idx="2">Sample</cx:pt></cx:lvl></cx:strDim><cx:strDim type="entityId"><cx:lvl ptCount="3"><cx:pt idx="2">sample</cx:pt></cx:lvl></cx:strDim><cx:numDim type="colorVal"><cx:lvl ptCount="3"><cx:pt idx="2">42</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="regionMap"><cx:dataId val="0"/><cx:layoutPr><cx:geography projectionType="${projection}"><cx:geoCache>${cache}</cx:geoCache></cx:geography></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
describe('shared region map painter', () => {
	it('projects cached polygons through the parser, summary and view while preserving point identity', () => {
		expect(writeChartExSpaceOverSource(xml(), parseChartExSpace(xml()).chartExSpace)).toBe(xml());
		const summary = chartExSummary(xml())!.summary!;
		expect(summary.chartType).toBe('regionMap');
		const svg = renderChartSummarySvg(summary, { width: 480, height: 300 });
		expect(svg).toContain('data-map-geometry="cached"');
		expect(svg).toContain('data-chart-point="2"');
		expect(svg).toContain('Sample: 42');
		expect(svg).toContain('Test attribution');
		expect(svg).not.toContain('Approximate country outlines');
		expect(svg).not.toContain('NaN');
	});
	it('keeps authored ChartEx title size over the native fallback', () => {
		const source = xml().replace(
			'<cx:plotArea>',
			'<cx:title><cx:tx><cx:rich xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="2400"/><a:t>Authored</a:t></a:r></a:p></cx:rich></cx:tx></cx:title><cx:plotArea>',
		);
		expect(chartExSummary(xml())!.summary!.formatting?.entries.title?.fontSize).toBe(18);
		expect(chartExSummary(source)!.summary!.formatting?.entries.title?.fontSize).toBe(24);
	});
	it('labels the fallback and does not fabricate polygons for unknown projections', () => {
		const fallback = renderChartSummarySvg(chartExSummary(xml(''))!.summary!, {
			width: 480,
			height: 300,
		});
		expect(fallback).toContain('Approximate country outlines');
		const unknown = renderChartSummarySvg(chartExSummary(xml(clear, 'unsupported'))!.summary!, {
			width: 480,
			height: 300,
		});
		expect(unknown).toContain('data-map-partial="true"');
		expect(unknown).not.toContain('<path');
		expect(unknown).toContain('Sample: 42');
	});
	it('honors authored solid series fill and owns one continuous legend', () => {
		const source = xml()
			.replace(
				'<cx:dataId',
				'<cx:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></cx:spPr><cx:dataId',
			)
			.replace('</cx:chart>', '<cx:legend pos="b"/></cx:chart>');
		const svg = renderChartSummarySvg(chartExSummary(source)!.summary!, {
			width: 480,
			height: 300,
		});
		expect(svg).toMatch(/<path[^>]*fill="#4472C4"/u);
		expect(svg).not.toContain('data-chart-part="legend"');
	});
});
