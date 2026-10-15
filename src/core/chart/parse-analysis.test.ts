import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import { chartTrendlineKind, chartTrendlineType } from './model-analysis';
import { parseChartSpace } from './parse-space';
import { writeChartSpace } from './write-space';

const SERIES_ANALYSIS =
	'<c:trendline><c:name>Fit &amp; trend</c:name><c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln></c:spPr>' +
	'<c:trendlineType val="poly"/><c:order val="3"/><c:forward val="0.5"/><c:backward val="1"/><c:intercept val="2"/><c:dispRSqr val="1"/><c:dispEq val="0"/>' +
	'<c:trendlineLbl><c:numFmt formatCode="0.00" sourceLinked="0"/></c:trendlineLbl></c:trendline>' +
	'<c:trendline><c:trendlineType val="movingAvg"/><c:period val="2"/></c:trendline>' +
	'<c:errBars><c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="cust"/><c:noEndCap val="1"/>' +
	'<c:plus><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:plus>' +
	'<c:minus><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>0.5</c:v></c:pt></c:numCache></c:numRef></c:minus></c:errBars>' +
	'<c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="percentage"/><c:val val="5"/></c:errBars>';

const part = (analysis: string) =>
	`<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:chart><c:plotArea><c:scatterChart><c:scatterStyle val="lineMarker"/>` +
	`<c:ser><c:idx val="0"/><c:order val="0"/>${analysis}<c:xVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:xVal>` +
	'<c:yVal><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:yVal></c:ser>' +
	'<c:axId val="1"/><c:axId val="2"/></c:scatterChart></c:plotArea></c:chart></c:chartSpace>';

describe('trendlines and error bars in the chart model', () => {
	it('reads every field of c:trendline and c:errBars', () => {
		const { chartSpace, issues } = parseChartSpace(part(SERIES_ANALYSIS));
		expect(issues).toEqual([]);
		const series = chartSpace.plotArea.groups[0]?.series[0];
		expect(series?.trendlines?.map(({ spPr: _spPr, ...rest }) => rest)).toEqual([
			{
				name: 'Fit & trend',
				trendlineType: 'poly',
				order: 3,
				forward: 0.5,
				backward: 1,
				intercept: 2,
				displayRSquared: true,
				displayEquation: false,
				label: { numberFormat: { formatCode: '0.00', sourceLinked: false } },
			},
			{ trendlineType: 'movingAvg', period: 2 },
		]);
		expect(series?.trendlines?.[0]?.spPr?.line).toMatchObject({ widthEmu: 19050, dash: 'dash' });
		const [custom, percent] = series?.errorBars ?? [];
		expect(custom).toMatchObject({
			direction: 'y',
			barType: 'plus',
			valueType: 'cust',
			noEndCap: true,
			plus: { kind: 'numLit' },
			minus: { kind: 'numRef', formula: 'Sheet1!$C$2:$C$3' },
		});
		expect(percent).toEqual({ direction: 'x', barType: 'both', valueType: 'percentage', value: 5 });
	});

	it('writes them back in schema order, byte for byte', () => {
		const source = part(SERIES_ANALYSIS);
		const { chartSpace } = parseChartSpace(source);
		const written = writeChartSpace(chartSpace);
		expect(written).toContain(SERIES_ANALYSIS);
		expect(parseChartSpace(written).chartSpace).toEqual(chartSpace);
	});

	it('keeps an unreadable trendline in place and writes the schema defaults', () => {
		const { chartSpace, issues } = parseChartSpace(
			part('<c:trendline><c:trendlineType val="cubic"/></c:trendline><c:errBars/>'),
		);
		expect(issues.map((issue) => issue.code)).toEqual(['CHART_VALUE_INVALID']);
		const series = chartSpace.plotArea.groups[0]?.series[0];
		expect(series?.trendlines).toEqual([{}]);
		expect(series?.errorBars).toEqual([{}]);
		const written = writeChartSpace(chartSpace);
		expect(written).toContain(
			'<c:trendline><c:trendlineType val="linear"/></c:trendline><c:errBars><c:errBarType val="both"/><c:errValType val="fixedVal"/></c:errBars>',
		);
	});

	it('maps trendline types to readable kinds and back', () => {
		expect(chartTrendlineKind('exp')).toBe('exponential');
		expect(chartTrendlineKind('poly')).toBe('polynomial');
		expect(chartTrendlineKind('toString')).toBeUndefined();
		expect(chartTrendlineKind(undefined)).toBeUndefined();
		expect(chartTrendlineType('logarithmic')).toBe('log');
	});
});
