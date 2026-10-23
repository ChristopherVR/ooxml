import { describe, expect, it } from 'vitest';
import { parseChartExSpace } from './parse';
import { DECLARATION, FULL, NS } from './test-support/full-part';
import { chartExWriteOptionsOf, writeChartExSpace } from './write';

describe('parseChartExSpace and writeChartExSpace', () => {
	it('round-trips every modelled element byte for byte', () => {
		const { chartExSpace, issues } = parseChartExSpace(FULL);
		expect(issues).toEqual([]);
		expect(writeChartExSpace(chartExSpace)).toBe(FULL);
	});

	it('reads the chart data, series options and axes', () => {
		const { chartExSpace: space } = parseChartExSpace(FULL);
		expect(space.chartData.externalData).toEqual({ relId: 'rId3', autoUpdate: true });
		const [cat, val] = space.chartData.data[0]!.dimensions;
		expect(cat).toMatchObject({
			kind: 'strDim',
			type: 'cat',
			formula: { formula: 'Sheet1!$A$2:$B$4', direction: 'row' },
			nameFormula: { formula: 'Sheet1!$A$1' },
		});
		expect(cat!.levels[0]!.points).toEqual([
			{ index: 0, value: 'a' },
			{ index: 2, value: ' c ' },
		]);
		expect(val!.levels[0]).toMatchObject({ pointCount: 3, formatCode: '0.0' });
		expect(space.title?.tx?.rich?.text).toBe('Title');
		const [histogram, pareto, map, box, waterfall, treemap, aggregated] = space.plotArea.series;
		expect(histogram).toMatchObject({
			layoutId: 'clusteredColumn',
			hidden: false,
			tx: { formula: { formula: 'Sheet1!$C$1' }, value: 'Sales' },
			dataId: 0,
			axisIds: [0, 1],
			layoutProperties: {
				binning: { intervalClosed: 'l', underflow: 'auto', overflow: 9.5, binSize: 2.5 },
			},
		});
		expect(histogram!.spPr?.fill).toEqual({
			kind: 'solid',
			color: { kind: 'srgb', value: '4472C4', transforms: [] },
		});
		expect(histogram!.dataLabels).toMatchObject({
			position: 'outEnd',
			numberFormat: { formatCode: '#,##0', sourceLinked: false },
			visibility: { seriesName: false, categoryName: true, value: true },
			separator: '; ',
			labels: [{ index: 2, position: 'ctr', visibility: { seriesName: true } }],
			hidden: [0],
		});
		expect(pareto).toMatchObject({ layoutId: 'paretoLine', ownerIndex: 0, axisIds: [2] });
		expect(map!.layoutProperties?.geography).toMatchObject({
			projectionType: 'mercator',
			attribution: 'Powered by Bing',
		});
		expect(map!.layoutProperties?.geography?.geoCacheXml).toContain('<cx:binary>QUJD</cx:binary>');
		expect(box!.layoutProperties).toEqual({
			visibility: { meanLine: true, meanMarker: false, nonoutliers: false, outliers: true },
			statistics: { quartileMethod: 'exclusive' },
		});
		expect(waterfall!.layoutProperties).toEqual({
			visibility: { connectorLines: false },
			subtotals: [0, 3],
		});
		expect(treemap!.layoutProperties).toEqual({ parentLabelLayout: 'banner' });
		expect(aggregated!.layoutProperties).toEqual({ aggregation: true });
		const [category, value] = space.plotArea.axes;
		expect(category).toMatchObject({
			id: 0,
			scaling: { kind: 'cat', gapWidth: 'auto' },
			title: { tx: { value: 'Bins' } },
			majorTickMarks: { type: 'out' },
			tickLabels: {},
		});
		expect(value).toMatchObject({
			hidden: true,
			scaling: { kind: 'val', max: 100, min: 'auto', majorUnit: 20, minorUnit: 'auto' },
			units: { unit: 'thousands', label: { tx: { value: 'Thousands' } } },
			minorGridlines: {},
			numberFormat: { formatCode: '0%', sourceLinked: true },
		});
		expect(space.legend).toMatchObject({ position: 'r', align: 'min', overlay: true });
	});

	it('writes edited fields and keeps the rest', () => {
		const { chartExSpace: space } = parseChartExSpace(FULL);
		space.plotArea.series[0]!.spPr!.fill = {
			kind: 'solid',
			color: { kind: 'srgb', value: 'ABCDEF', transforms: [] },
		};
		space.plotArea.series[0]!.tx!.value = 'Revenue & cost';
		delete space.legend;
		const written = writeChartExSpace(space);
		expect(written).toContain('<a:srgbClr val="ABCDEF"/>');
		expect(written).toContain('<cx:v>Revenue &amp; cost</cx:v>');
		expect(written).not.toContain('<cx:legend');
		const again = parseChartExSpace(written).chartExSpace;
		expect(again.plotArea.series[0]!.spPr?.fill).toEqual(space.plotArea.series[0]!.spPr?.fill);
		expect(again.plotArea.series[0]!.tx).toEqual(space.plotArea.series[0]!.tx);
		expect(again.plotArea.axes).toEqual(space.plotArea.axes);
		expect(again.chartData).toEqual(space.chartData);
	});

	it('reports unknown children, malformed values and a classic root', () => {
		const part = `<cx:chartSpace ${NS}><cx:chartData><cx:data id="x"/></cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:mystery/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`;
		const { chartExSpace, issues } = parseChartExSpace(part);
		expect(issues.map((issue) => issue.code)).toEqual([
			'CHART_VALUE_INVALID',
			'CHART_ELEMENT_NOT_MODELLED',
		]);
		expect(chartExSpace.plotArea.series[0]!.layoutId).toBe('funnel');
		expect(writeChartExSpace(chartExSpace)).not.toContain('mystery');
		const classic = parseChartExSpace(
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>',
		);
		expect(classic.issues.map((issue) => issue.code)).toEqual(['CHART_ROOT_UNEXPECTED']);
	});

	it('writes the default declarations, user shapes and the empty-element style', () => {
		const written = writeChartExSpace(
			{
				chartData: { data: [] },
				plotArea: { series: [{ layoutId: 'funnel', dataPoints: [], axisIds: [] }], axes: [] },
				userShapesRelId: 'rId2',
			},
			{ emptyElementSpace: true },
		);
		expect(written).toBe(
			`${DECLARATION}<cx:chartSpace ${NS} xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><cx:chartData /><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel" /></cx:plotAreaRegion></cx:plotArea></cx:chart><c:userShapes r:id="rId2" /></cx:chartSpace>`,
		);
		expect(parseChartExSpace(written).chartExSpace.userShapesRelId).toBe('rId2');
	});
});

describe('chartExWriteOptionsOf', () => {
	it('detects the declaration break and a space before an empty element', () => {
		expect(chartExWriteOptionsOf('<?xml version="1.0"?>\r\n<a><b /></a>')).toEqual({
			declarationBreak: '\r\n',
			emptyElementSpace: true,
		});
		expect(chartExWriteOptionsOf('<?xml version="1.0"?><a><b/></a>')).toEqual({
			declarationBreak: '',
		});
	});

	it('scans a long run of unclosed tags without backtracking', () => {
		const start = performance.now();
		expect(chartExWriteOptionsOf('<'.repeat(50_000))).toEqual({ declarationBreak: '' });
		expect(performance.now() - start).toBeLessThan(1000);
	});
});
