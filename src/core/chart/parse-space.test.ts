import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { chartCacheValues, chartSourceValues } from './data-cache';
import { parseChartSpace } from './parse-space';

const space = (body: string, head = '') =>
	`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">${head}<c:chart>${body}</c:chart></c:chartSpace>`;

describe('parseChartSpace', () => {
	it('reads the part-level settings, the Office style fallback and raw extension lists', () => {
		const { chartSpace, issues } = parseChartSpace(
			space(
				'<c:autoTitleDeleted/><c:plotArea/><c:plotVisOnly val="0"/><c:dispBlanksAs val="span"/><c:extLst><c:ext uri="{X}"/></c:extLst>',
				'<c:date1904 val="1"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent><mc:Choice Requires="c14"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent>',
			),
		);
		expect(issues).toEqual([]);
		expect(chartSpace).toMatchObject({
			date1904: true,
			language: 'en-US',
			roundedCorners: false,
			style: 2,
			// CT_Boolean without @val means true.
			autoTitleDeleted: true,
			plotVisibleOnly: false,
			displayBlanksAs: 'span',
		});
		expect(chartSpace.chartExtLst).toContain('uri="{X}"');
	});

	it('reads series details: points, labels, markers, literals and multi-level categories', () => {
		const ser = `<c:ser><c:idx val="3"/><c:order val="1"/><c:tx><c:v>Literal</c:v></c:tx>
			<c:spPr><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill><a:ln w="12700"><a:noFill/></a:ln><a:effectLst><a:glow rad="1"/></a:effectLst></c:spPr>
			<c:invertIfNegative val="1"/><c:marker><c:symbol val="diamond"/><c:size val="7"/></c:marker>
			<c:dPt><c:idx val="1"/><c:explosion val="12"/><c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr></c:dPt>
			<c:dLbls><c:dLbl><c:idx val="1"/><c:tx><c:rich><a:p><a:r><a:t>Peak</a:t></a:r></a:p></c:rich></c:tx><c:showVal val="1"/></c:dLbl><c:numFmt formatCode="0%" sourceLinked="0"/><c:dLblPos val="outEnd"/><c:showVal val="1"/><c:showLeaderLines val="0"/></c:dLbls>
			<c:cat><c:multiLvlStrRef><c:f>S!$A$1:$B$2</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>Group</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
			<c:val><c:numLit><c:formatCode>0.0</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1.5</c:v></c:pt><c:pt idx="2"><c:v>x</c:v></c:pt></c:numLit></c:val>
			<c:smooth/></c:ser>`;
		const { chartSpace, issues } = parseChartSpace(
			space(
				`<c:plotArea><c:lineChart><c:grouping val="standard"/>${ser}<c:marker val="1"/><c:smooth val="0"/><c:axId val="5"/></c:lineChart></c:plotArea>`,
			),
		);
		expect(issues).toEqual([]);
		// CT_LineChart's own `c:smooth` is the group default; the series keeps its override.
		expect(chartSpace.plotArea.groups[0]).toMatchObject({ marker: true, smooth: false });
		const series = chartSpace.plotArea.groups[0]?.series[0];
		expect(series).toMatchObject({ index: 3, order: 1, invertIfNegative: true, smooth: true });
		expect(series?.tx).toEqual({ value: 'Literal', text: 'Literal' });
		expect(series?.spPr?.fill).toEqual({
			kind: 'solid',
			color: { kind: 'srgb', value: 'ff0000', transforms: [] },
		});
		expect(series?.spPr?.line).toEqual({ widthEmu: 12700, fill: { kind: 'none' } });
		expect(series?.spPr?.effectsXml).toContain('a:glow');
		expect(series?.marker).toEqual({ symbol: 'diamond', size: 7 });
		expect(series?.dataPoints).toEqual([
			{
				index: 1,
				explosion: 12,
				spPr: {
					fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2', transforms: [] } },
					sourceXml: '<a:solidFill><a:schemeClr val="accent2"/></a:solidFill>',
				},
			},
		]);
		expect(series?.dataLabels).toMatchObject({
			showValue: true,
			showLeaderLines: false,
			position: 'outEnd',
			numberFormat: { formatCode: '0%', sourceLinked: false },
			labels: [{ index: 1, showValue: true, tx: { text: 'Peak' } }],
		});
		expect(series?.categories?.levels).toHaveLength(2);
		expect(chartSourceValues(series?.categories)).toEqual(['a', 'b']);
		expect(series?.values).toMatchObject({
			kind: 'numLit',
			cache: { formatCode: '0.0', pointCount: 3 },
		});
		expect(chartCacheValues(series?.values?.cache)).toEqual([1.5, '', 'x']);
	});

	it('reads axes: scaling, titles, number formats, ticks and crossing', () => {
		const { chartSpace } = parseChartSpace(
			space(`<c:plotArea><c:barChart><c:barDir val="bar"/><c:axId val="1"/><c:axId val="2"/></c:barChart>
				<c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="l"/><c:numFmt formatCode="d-mmm" sourceLinked="1"/><c:crossAx val="2"/><c:baseTimeUnit val="days"/><c:auto val="0"/></c:dateAx>
				<c:valAx><c:axId val="2"/><c:scaling><c:logBase val="10"/><c:orientation val="maxMin"/><c:max val="1000"/><c:min val="1"/></c:scaling><c:axPos val="b"/><c:majorGridlines/><c:minorGridlines/>
				<c:title><c:tx><c:rich><a:p><a:r><a:t>Amount</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
				<c:majorTickMark val="cross"/><c:minorTickMark val="in"/><c:tickLblPos val="high"/><c:crossAx val="1"/><c:crossesAt val="10"/><c:crossBetween val="between"/><c:majorUnit val="100"/></c:valAx></c:plotArea>
				<c:legend><c:legendPos val="tr"/><c:legendEntry><c:idx val="0"/><c:delete val="1"/></c:legendEntry><c:overlay val="1"/></c:legend>`),
		);
		expect(chartSpace.plotArea.groups[0]?.barDirection).toBe('bar');
		const [date, value] = chartSpace.plotArea.axes;
		expect(date).toMatchObject({
			kind: 'date',
			id: 1,
			crossAxisId: 2,
			deleted: true,
			position: 'l',
			numberFormat: { formatCode: 'd-mmm', sourceLinked: true },
			baseTimeUnit: 'days',
			auto: false,
			majorGridlines: false,
		});
		expect(value).toMatchObject({
			kind: 'val',
			scaling: { orientation: 'maxMin', logBase: 10, min: 1, max: 1000 },
			majorGridlines: true,
			minorGridlines: true,
			title: { text: 'Amount', overlay: false },
			majorTickMark: 'cross',
			minorTickMark: 'in',
			tickLabelPosition: 'high',
			crossesAt: 10,
			crossBetween: 'between',
			majorUnit: 100,
		});
		expect(chartSpace.legend).toEqual({
			position: 'tr',
			overlay: true,
			entries: [{ index: 0, deleted: true }],
		});
	});

	it('reports unknown children and malformed values instead of throwing', () => {
		const { chartSpace, issues } = parseChartSpace(
			space(
				'<c:plotArea><c:barChart><c:gapWidth val="wide"/><c:mystery/><c:ser><c:errBars/></c:ser></c:barChart><c:dTable/></c:plotArea>',
				'<c:protection/>',
			),
		);
		expect(chartSpace.plotArea.groups[0]?.gapWidth).toBeUndefined();
		expect(issues.map((issue) => issue.message)).toEqual([
			'c:chartSpace/c:protection is not modelled; it is kept only in the source part.',
			'c:barChart/c:mystery is not modelled; it is kept only in the source part.',
			'c:ser/c:errBars is not modelled; it is kept only in the source part.',
			'c:barChart/c:gapWidth has an unreadable value "wide".',
			'c:plotArea/c:dTable is not modelled; it is kept only in the source part.',
		]);
	});

	it('accepts an element, and reports parts that are not classic charts', () => {
		const element = parseXml(
			space('<c:plotArea><c:ofPieChart><c:ofPieType val="bar"/></c:ofPieChart></c:plotArea>'),
		).documentElement;
		expect(parseChartSpace(element).chartSpace.plotArea.groups[0]).toMatchObject({
			kind: 'ofPie',
			ofPieType: 'bar',
		});
		const chartex = parseChartSpace(
			'<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>',
		);
		expect(chartex.issues.map((issue) => issue.code)).toEqual(['CHART_ROOT_UNEXPECTED']);
		const empty = parseChartSpace(
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
		);
		expect(empty.issues.map((issue) => issue.code)).toEqual([
			'CHART_CHART_MISSING',
			'CHART_PLOT_AREA_MISSING',
		]);
	});
});
