import { describe, expect, it } from 'vitest';

import {
	readChartElementFormat,
	withChartElementFormat,
} from '../../builders/sdk/chart-element-format-operations';
import { withChartDataLabelOptions } from '../../builders/sdk/chart-format-options-operations';
import {
	isChartSeriesOnSecondaryAxis,
	withChartSeriesOnSecondaryAxis,
} from '../../builders/sdk/chart-series-axis-operations';
import { BAR_CHART_XML, saveWith, section } from './chart-element-format-deck.test-util';

/**
 * Round 6 of the chart Format pane: picture, pattern and gradient fills,
 * gradient lines, 3-D Format, Text Options, Plot Series on Secondary Axis and
 * per-series data labels, each saved into the right element and read back.
 */

const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

/** The fixture chart with a second series (West) in the same bar group. */
const TWO_SERIES_XML = BAR_CHART_XML.replace(
	'<c:gapWidth val="150"/>',
	`<c:ser>
		<c:idx val="1"/>
		<c:order val="1"/>
		<c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>West</c:v></c:pt></c:strCache></c:strRef></c:tx>
		<c:invertIfNegative val="0"/>
		<c:cat><c:strRef><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat>
		<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>300</c:v></c:pt><c:pt idx="1"><c:v>400</c:v></c:pt></c:numCache></c:numRef></c:val>
	</c:ser>
	<c:gapWidth val="150"/>`,
);

describe('chart Format pane fills, lines, 3-D and text save and reload', () => {
	it('writes pattern, gradient and picture fills and reads them back', async () => {
		const { xml, zip, reloaded } = await saveWith((data) => {
			let next = withChartElementFormat(
				data,
				{ kind: 'plotArea' },
				{
					fill: {
						kind: 'pattern',
						pattern: { preset: 'ltUpDiag', foreground: '#C00000', background: '#FFFFFF' },
					},
				},
			);
			next = withChartElementFormat(
				next,
				{ kind: 'legend' },
				{
					fill: {
						kind: 'gradient',
						gradient: {
							type: 'linear',
							angle: 0,
							stops: [
								{ color: '#FFFFFF', position: 0 },
								{ color: '#5B9BD5', position: 100 },
							],
						},
					},
				},
			);
			return withChartElementFormat(
				next,
				{ kind: 'chartArea' },
				{ fill: { kind: 'picture', picture: { src: PNG, transparency: 20 } } },
			);
		});
		expect(section(xml, 'c:plotArea')).toMatch(
			/<c:spPr><a:pattFill prst="ltUpDiag"><a:fgClr><a:srgbClr val="C00000"\/><\/a:fgClr><a:bgClr><a:srgbClr val="FFFFFF"\/>/u,
		);
		expect(section(xml, 'c:legend')).toMatch(/<c:spPr><a:gradFill rotWithShape="1"><a:gsLst>/u);
		const chartSpaceSpPr = /<\/c:chart>\s*<c:spPr>([\s\S]*?)<\/c:spPr>/u.exec(xml)?.[1] ?? '';
		const embed =
			/<a:blipFill rotWithShape="1"><a:blip r:embed="(rId\d+)"><a:alphaModFix amt="80000"\/><\/a:blip><a:stretch><a:fillRect\/><\/a:stretch><\/a:blipFill>/u.exec(
				chartSpaceSpPr,
			)?.[1];
		expect(embed).toBeDefined();
		const rels = await zip.file('ppt/charts/_rels/chart1.xml.rels')!.async('string');
		const target = new RegExp(
			`Id="${embed}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="${embed}"`,
			'u',
		).exec(rels);
		const media = (target?.[1] ?? target?.[2] ?? '').replace(/\.\.\//gu, 'ppt/');
		expect(media).toMatch(/^ppt\/media\/image\d+\.png$/u);
		expect(zip.file(media)).not.toBeNull();
		expect(await zip.file('[Content_Types].xml')!.async('string')).toMatch(/Extension="png"/u);

		const data = reloaded.chart.chartData!;
		expect(readChartElementFormat(data, { kind: 'plotArea' }).fill).toMatchObject({
			kind: 'pattern',
			pattern: { preset: 'ltUpDiag' },
		});
		expect(readChartElementFormat(data, { kind: 'legend' }).fill.kind).toBe('gradient');
		const area = readChartElementFormat(data, { kind: 'chartArea' }).fill;
		expect(area.kind).toBe('picture');
		expect(area.picture?.transparency).toBe(20);
		expect(area.picture?.src).toMatch(/^data:image\/png/u);
	});

	it('writes a gradient line, 3-D Format and Text Options', async () => {
		const { xml, reloaded } = await saveWith((data) => {
			let next = withChartElementFormat(
				data,
				{ kind: 'axis', axisIndex: 0 },
				{
					line: {
						kind: 'gradient',
						gradient: {
							type: 'linear',
							angle: 0,
							stops: [
								{ color: '#000000', position: 0 },
								{ color: '#FF0000', position: 100 },
							],
						},
					},
					text: { fill: { kind: 'solid', color: '#C00000' }, rotation: -45 },
				},
			);
			next = withChartElementFormat(
				next,
				{ kind: 'series', seriesIndex: 0 },
				{
					threeD: {
						bevelTop: { type: 'circle', width: 6, height: 6 },
						material: 'metal',
						lighting: 'threePt',
					},
				},
			);
			return withChartElementFormat(
				next,
				{ kind: 'title' },
				{
					text: {
						fill: { kind: 'solid', color: '#2F5597', transparency: 10 },
						outline: { kind: 'solid', color: '#000000', width: 0.5 },
						effects: { glow: { color: '#FFC000', radius: 4, transparency: 40 } },
					},
				},
			);
		});
		const catAx = section(xml, 'c:catAx');
		expect(catAx).toMatch(/<c:spPr><a:ln><a:gradFill rotWithShape="1">/u);
		expect(catAx).toMatch(
			/<\/c:spPr>\s*<c:txPr><a:bodyPr rot="-2700000"\/><a:lstStyle\/><a:p><a:pPr><a:defRPr><a:solidFill><a:srgbClr val="C00000"\/>/u,
		);
		expect(section(xml, 'c:ser')).toMatch(
			/<a:scene3d><a:camera prst="orthographicFront"\/><a:lightRig dir="t" rig="threePt"\/><\/a:scene3d><a:sp3d prstMaterial="metal"><a:bevelT w="76200" h="76200" prst="circle"\/><\/a:sp3d>/u,
		);
		// The title's own run carries the text fill, outline and glow (outline first, CT_TextCharacterProperties order).
		expect(section(xml, 'c:title')).toMatch(
			/<a:r><a:rPr><a:ln w="6350"><a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill><\/a:ln><a:solidFill><a:srgbClr val="2F5597"><a:alpha val="90000"\/><\/a:srgbClr><\/a:solidFill><a:effectLst><a:glow rad="50800">/u,
		);

		const data = reloaded.chart.chartData!;
		expect(readChartElementFormat(data, { kind: 'axis', axisIndex: 0 }).line.kind).toBe('gradient');
		expect(readChartElementFormat(data, { kind: 'axis', axisIndex: 0 }).text).toMatchObject({
			rotation: -45,
			fill: { kind: 'solid' },
		});
		expect(readChartElementFormat(data, { kind: 'series', seriesIndex: 0 }).threeD).toMatchObject({
			bevelTop: { type: 'circle', width: 6, height: 6 },
			material: 'metal',
			lighting: 'threePt',
		});
		const title = readChartElementFormat(data, { kind: 'title' }).text;
		expect(title.fill).toMatchObject({ kind: 'solid', transparency: 10 });
		expect(title.outline).toMatchObject({ kind: 'solid', width: 0.5 });
		expect(title.effects?.glow).toMatchObject({ radius: 4, transparency: 40 });
	});
});

describe('Plot Series on Secondary Axis', () => {
	it('creates the secondary axes and a chart group of its own, and removes them again', async () => {
		const { xml, reloaded } = await saveWith(
			(data) => withChartSeriesOnSecondaryAxis(data, 1, true),
			TWO_SERIES_XML,
		);
		const groups = xml.match(/<c:barChart>[\s\S]*?<\/c:barChart>/gu) ?? [];
		expect(groups).toHaveLength(2);
		expect(groups[0]).toContain('<c:v>East</c:v>');
		expect(groups[0]).not.toContain('<c:v>West</c:v>');
		expect(groups[1]).toContain('<c:v>West</c:v>');
		expect(groups[1]).toMatch(/<c:axId val="224"\/>\s*<c:axId val="223"\/>/u);
		expect(xml).toMatch(
			/<c:valAx><c:axId val="223"\/><c:scaling><c:orientation val="minMax"\/><\/c:scaling><c:delete val="0"\/><c:axPos val="r"\/>[\s\S]*?<c:crossAx val="224"\/><c:crosses val="max"\/>/u,
		);
		expect(xml).toMatch(/<c:catAx><c:axId val="224"\/>[\s\S]*?<c:delete val="1"\/>/u);

		const data = reloaded.chart.chartData!;
		expect(isChartSeriesOnSecondaryAxis(data, 1)).toBe(true);
		expect(isChartSeriesOnSecondaryAxis(data, 0)).toBe(false);

		// Back to the primary axis: one group again and no orphaned secondary axes.
		const back = await saveWith((d) => withChartSeriesOnSecondaryAxis(d, 1, false), TWO_SERIES_XML);
		expect(back.xml.match(/<c:barChart>/gu)).toHaveLength(1);
		expect(back.xml).not.toContain('<c:axPos val="r"/>');
	});
});

describe('per-series data label options', () => {
	it('shows labels on one series only and writes that series its own c:dLbls', async () => {
		const { xml, reloaded } = await saveWith(
			(data) => withChartDataLabelOptions(data, { showValue: true, showSeriesName: true }, 1),
			TWO_SERIES_XML,
		);
		const series = xml.match(/<c:ser>[\s\S]*?<\/c:ser>/gu) ?? [];
		expect(series[1]).toMatch(
			/<c:dLbls><c:showLegendKey val="0"\/><c:showVal val="1"\/><c:showCatName val="0"\/><c:showSerName val="1"\/>/u,
		);
		// The first series keeps its own authored labels untouched.
		expect(series[0]).toContain(
			'<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>',
		);
		const data = reloaded.chart.chartData!;
		expect(data.series[1].dataLabelOptions).toMatchObject({
			showValue: true,
			showSeriesName: true,
		});
	});
});
