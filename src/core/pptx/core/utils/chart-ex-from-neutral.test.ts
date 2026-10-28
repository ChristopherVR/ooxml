import { XMLParser } from 'fast-xml-parser';
import { beforeAll, describe, expect, it } from 'vitest';
import { parseChartExSpace } from '../../../chart';
import { CHART_EX_FIXTURES, chartExParts } from '../../../chart/ex/test-support/fixtures';
import { FULL } from '../../../chart/ex/test-support/full-part';
import { PptxXmlLookupService } from '../services/PptxXmlLookupService';
import { parseCxChartSeries } from './chart-cx-parser';
import { chartExFromNeutral } from './chart-ex-from-neutral';
import { readChartExStyleBridge } from './chart-ex-style-bridge';
import type { XmlObject } from '../types';
const parts: [string, string][] = [];
beforeAll(async () => {
	for (const [file] of CHART_EX_FIXTURES) parts.push(...(await chartExParts(file)));
	parts.push(['synthetic', FULL]);
});
const lookup = new PptxXmlLookupService();
const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false, trimValues: false });
function compare(xml: string) {
	const document = parser.parse(xml) as XmlObject;
	const root = lookup.getChildByLocalName(document, 'chartSpace')!;
	const chart = lookup.getChildByLocalName(root, 'chart')!;
	const plot = lookup.getChildByLocalName(chart, 'plotArea')!;
	const color = { parseColor: (node: XmlObject | undefined) => (node ? '#123456' : undefined) };
	const fonts: string[] = [];
	const font = (raw: string) => {
		fonts.push(raw);
		return `Resolved ${raw}`;
	};
	const old = parseCxChartSeries(plot, lookup, root, chart, color, font)!;
	const neutral = chartExFromNeutral(
		parseChartExSpace(xml).chartExSpace,
		readChartExStyleBridge(plot, lookup, root, color, font),
	);
	return { old, neutral, fonts };
}
describe('ChartEx neutral read adapter', () => {
	it('projects numeric funnel category gaps, preserving explicit zero and native auto defaults', () => {
		const space = parseChartExSpace(FULL).chartExSpace;
		space.plotArea.series[0]!.layoutId = 'funnel';
		const styled = { categories: [], series: [] };
		for (const value of [0, 100]) {
			space.plotArea.axes = [{ id: 0, scaling: { kind: 'cat', gapWidth: value } }];
			expect(chartExFromNeutral(space, styled).chartData?.funnelGapWidth).toBe(value);
		}
		space.plotArea.axes = [{ id: 0, scaling: { kind: 'cat', gapWidth: 'auto' } }];
		expect(chartExFromNeutral(space, styled).chartData?.funnelGapWidth).toBeUndefined();
	});
	it('retains the complete PowerPoint view and host-resolved styles on every committed part', () => {
		for (const [name, xml] of parts) {
			const { neutral, old } = compare(xml);
			// The old reader omitted group content switches and numeric axis scaling. They are tested
			// against authored neutral facts below; every previously exposed field still compares here.
			for (const series of neutral.series)
				if (series.dataLabelOptions) {
					for (const key of ['showValue', 'showCategory', 'showSeriesName'] as const)
						delete series.dataLabelOptions[key];
					if (!Object.keys(series.dataLabelOptions).length) delete series.dataLabelOptions;
				}
			for (const axis of neutral.chartData?.axes ?? [])
				for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) delete axis[key];
			for (const [index, series] of neutral.series.entries()) {
				// The map adapter additionally exposes explicit fill and source-index dimensions.
				// Their semantics and sparse editing round-trip are covered by dedicated regressions.
				if (series.regionMapOptions) {
					if (series.regionMapOptions.fixedColor)
						expect(series.regionMapOptions.fixedColor).toBe(series.color);
					delete series.regionMapOptions.fixedColor;
					delete series.regionMapOptions.geographyCacheXml;
					for (const key of [
						'categorySourceIndices',
						'entityIdSourceIndices',
						'valueSourceIndices',
					] as const)
						if (!old.series[index]?.regionMapOptions?.[key]) delete series.regionMapOptions[key];
				}
				series.dataLabels = series.dataLabels?.filter(
					(label) =>
						!label.deleted || old.series[index]?.dataLabels?.some((item) => item.idx === label.idx),
				);
				for (const label of series.dataLabels ?? []) delete label.deleted;
				if (!series.dataLabels?.length) delete series.dataLabels;
			}
			expect(neutral, name).toEqual(old);
		}
		expect(parts).toHaveLength(16);
	});
	it('reads group label switches, point hiding and numeric axis bounds omitted by the legacy reader', () => {
		const model = parseChartExSpace(FULL).chartExSpace;
		const axis = model.plotArea.axes.find((axis) => axis.scaling?.kind === 'val')!;
		axis.scaling = { kind: 'val', min: -2, max: 40, majorUnit: 5 };
		const group = model.plotArea.series[0]!.dataLabels!;
		group.visibility = { value: false, categoryName: true };
		group.labels = [{ index: 0, visibility: { value: true } }];
		group.hidden = [0];
		const result = chartExFromNeutral(model, { categories: [], series: [] });
		expect(result.series[0]!.dataLabelOptions).toMatchObject({
			showValue: false,
			showCategory: true,
		});
		expect(result.series[0]!.dataLabels![0]!.deleted).toBe(true);
		expect(result.chartData!.axes!.find((item) => item.axisType === 'valAx')).toMatchObject({
			min: -2,
			max: 40,
			majorUnit: 5,
		});
	});
	it('preserves host font and colour resolution while neutral facts supply the data', () => {
		const { neutral, fonts } = compare(
			FULL.replace(
				'<a:defRPr sz="900"/>',
				'<a:defRPr sz="900"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr>',
			),
		);
		expect(fonts.length).toBeGreaterThan(0);
		expect(neutral.series.some((series) => series.color === '#123456')).toBe(true);
		expect(neutral.chartData?.axes?.some((axis) => axis.fontFamily?.startsWith('Resolved '))).toBe(
			true,
		);
	});
});
