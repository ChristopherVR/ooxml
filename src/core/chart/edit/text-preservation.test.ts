import { describe, expect, it } from 'vitest';
import { columnChart, COLUMN_CHART_XML } from './test-support/column-chart';
import { parseChartSpace } from '../parse-space';
import { writeChartSpace } from '../write-space';
import { applyChartAppearance } from './apply-appearance';
import { setChartPointLabel } from './point-label';
import { setChartDisplayUnitsLabel } from './axis-label';

const outline = (token: string) =>
	`<a:ln w="12700"><a:solidFill><a:schemeClr val="${token}"><a:lumMod val="60000"/></a:schemeClr></a:solidFill><a:round/><a:tailEnd type="arrow" w="lg"/><a:extLst><a:ext uri="line-private"/></a:extLst></a:ln>`;
const gradient = (token: string) =>
	`<a:gradFill flip="x"><a:gsLst><a:gs pos="0"><a:schemeClr val="${token}"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent3"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/><a:extLst><a:ext uri="gradient-private"/></a:extLst></a:gradFill>`;
const shadow = (token: string) =>
	`<a:effectLst><a:outerShdw blurRad="12700" dist="12700" dir="0"><a:schemeClr val="${token}"/></a:outerShdw><a:reflection blurRad="0"/></a:effectLst>`;
function titleWith(paint: (token: string) => string) {
	return parseChartSpace(
		COLUMN_CHART_XML.replace(
			'<a:rPr lang="en-US" sz="1400" b="1"/>',
			`<a:rPr lang="en-US" sz="1400" b="1">${paint('accent1')}</a:rPr>`,
		).replace('<a:rPr lang="en-US"/>', `<a:rPr lang="en-US">${paint('accent2')}</a:rPr>`),
	).chartSpace;
}
const rich = (space: ReturnType<typeof columnChart>) =>
	parseChartSpace(writeChartSpace(space)).chartSpace.title!.tx!.rich!.sourceXml!;

describe('resolved text edits preserve each authored rich run', () => {
	it('changes outline width and opacity without pinning colors or clearing joins, arrows and extensions', () => {
		const source = titleWith(outline);
		const line = { kind: 'solid' as const, color: '#123456', width: 1, transparency: 0 };
		const changed = applyChartAppearance(
			source,
			{ kind: 'title' },
			{ text: { outline: line } },
			{ text: { outline: { ...line, width: 3, transparency: 25 } } },
		);
		const xml = rich(changed);
		for (const token of ['accent1', 'accent2']) expect(xml).toContain(`val="${token}"`);
		expect(
			[...xml.matchAll(/<a:rPr\b[\s\S]*?<\/a:rPr>/g)].map((match) => match[0]).join(''),
		).not.toContain('srgbClr');
		expect(xml).toContain('w="38100"');
		expect(xml).toContain('alpha val="75000"');
		expect(xml).toContain('lumMod val="60000"');
		expect(xml).toContain('line-private');
		expect(xml).toContain('<a:round/>');
		expect(xml).toContain('type="arrow"');
		expect(rich(source)).toContain('w="12700"');
	});
	it('changes gradient geometry without flattening distinct run themes or producer extensions', () => {
		const source = titleWith(gradient);
		const fill = {
			kind: 'gradient' as const,
			gradient: {
				type: 'linear' as const,
				angle: 0,
				scaled: false,
				stops: [
					{ position: 0, color: '#123456' },
					{ position: 100, color: '#ABCDEF' },
				],
			},
		};
		const xml = rich(
			applyChartAppearance(
				source,
				{ kind: 'title' },
				{ text: { fill } },
				{ text: { fill: { ...fill, gradient: { ...fill.gradient, angle: 45 } } } },
			),
		);
		for (const token of ['accent1', 'accent2', 'accent3']) expect(xml).toContain(`val="${token}"`);
		expect(
			[...xml.matchAll(/<a:rPr\b[\s\S]*?<\/a:rPr>/g)].map((match) => match[0]).join(''),
		).not.toContain('srgbClr');
		expect(xml).toContain('ang="2700000"');
		expect(xml).toContain('gradient-private');
		expect(xml).toContain('flip="x"');
	});
	it('changes text shadow blur without pinning run colors or removing unrelated effects', () => {
		const source = titleWith(shadow);
		const effects = { shadow: { color: '#123456', blur: 1, distance: 1, angle: 0 } };
		const xml = rich(
			applyChartAppearance(
				source,
				{ kind: 'title' },
				{ text: { effects } },
				{ text: { effects: { shadow: { ...effects.shadow, blur: 4 } } } },
			),
		);
		expect(xml).toContain('val="accent1"');
		expect(xml).toContain('val="accent2"');
		expect(
			[...xml.matchAll(/<a:rPr\b[\s\S]*?<\/a:rPr>/g)].map((match) => match[0]).join(''),
		).not.toContain('srgbClr');
		expect(xml).toContain('blurRad="50800"');
		expect(xml).toContain('<a:reflection');
	});
	it('formats rich point and display-unit labels as well as their default properties', () => {
		const source = columnChart(),
			body = structuredClone(source.title!.tx!.rich!);
		source.plotArea.groups[0]!.series[0]!.dataLabels = {
			labels: [
				{
					index: 2,
					tx: { rich: body, text: 'Sales 2026' },
					extLst: '<c:extLst><c:ext uri="point-private"/></c:extLst>',
				},
			],
		};
		source.plotArea.axes[1]!.displayUnits = {
			builtInUnit: 'thousands',
			extLst: '<c:extLst><c:ext uri="units-private"/></c:extLst>',
			label: { tx: { rich: body, text: 'Sales 2026' } },
		};
		const points = setChartPointLabel(source, 0, 2, { font: { sizePt: 22 } });
		const changed = setChartDisplayUnitsLabel(points, { index: 1 }, { font: { sizePt: 24 } });
		const saved = parseChartSpace(writeChartSpace(changed)).chartSpace;
		expect(
			saved.plotArea.groups[0]!.series[0]!.dataLabels!.labels[0]!.tx!.rich!.sourceXml,
		).toContain('sz="2200"');
		expect(saved.plotArea.axes[1]!.displayUnits!.label!.tx!.rich!.sourceXml).toContain('sz="2400"');
		expect(writeChartSpace(changed)).toContain('point-private');
		expect(writeChartSpace(changed)).toContain('units-private');
		expect(body.sourceXml).toContain('sz="1400"');
	});
});
