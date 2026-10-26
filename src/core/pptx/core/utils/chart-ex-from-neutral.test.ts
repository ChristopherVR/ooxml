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
	it('retains the complete PowerPoint view and host-resolved styles on every committed part', () => {
		for (const [name, xml] of parts) expect(compare(xml).neutral, name).toEqual(compare(xml).old);
		expect(parts).toHaveLength(16);
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
