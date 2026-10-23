import { describe, expect, it } from 'vitest';
import type { DrawingFill } from '../../drawingml/types';
import { parseChartExSpace } from '../ex/parse';
import { FULL } from '../ex/test-support/full-part';
import { chartExWriteOptionsOf, writeChartExSpace } from '../ex/write';
import {
	renameChartExSeries,
	setChartExBinning,
	setChartExBoxWhisker,
	setChartExDataLabelsVisible,
	setChartExLegend,
	setChartExParentLabelLayout,
	setChartExSeriesFill,
	setChartExTitle,
	setChartExWaterfall,
} from './chartex';
import { setChartExCategories, setChartExSeriesValues } from './chartex-data';

// Word's waterfall part (docx/__fixtures__/chart-families.docx, word/charts/chartEx1.xml), shortened.
const WATERFALL =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:externalData r:id="rId1" cx:autoUpdate="0" /><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$E$2:$E$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Start</cx:pt><cx:pt idx="1">Sales</cx:pt><cx:pt idx="2">Total</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>Sheet1!$F$2:$F$4</cx:f><cx:lvl ptCount="3" formatCode="0.0"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">60</cx:pt><cx:pt idx="2">160</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Cash flow</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall" uniqueId="{X}"><cx:tx><cx:txData><cx:f>Sheet1!$F$1</cx:f><cx:v>Amount</cx:v></cx:txData></cx:tx><cx:dataId val="0" /><cx:layoutPr><cx:visibility connectorLines="1" /><cx:subtotals><cx:idx val="0" /><cx:idx val="2" /></cx:subtotals></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis></cx:plotArea><cx:legend pos="t" align="ctr" overlay="0" /></cx:chart></cx:chartSpace>';

const space = parseChartExSpace(WATERFALL).chartExSpace;
const write = (edited: typeof space) => writeChartExSpace(edited, chartExWriteOptionsOf(WATERFALL));

describe('chartex edit operations', () => {
	it('renames a series, keeping its cell reference and every other byte', () => {
		expect(write(renameChartExSeries(space, 0, 'Net'))).toBe(
			WATERFALL.replace('<cx:v>Amount</cx:v>', '<cx:v>Net</cx:v>'),
		);
		const unlinked = renameChartExSeries(space, 0, 'Net', { keepReference: false });
		expect(unlinked.plotArea.series[0]!.tx).toEqual({ value: 'Net' });
		expect(space.plotArea.series[0]!.tx?.value).toBe('Amount');
	});

	it('sets values and categories into the referenced data, keeping formulas and format codes', () => {
		const edited = setChartExCategories(setChartExSeriesValues(space, 0, [10, null, 30]), [
			'A',
			'B',
			'C',
		]);
		const written = write(edited);
		expect(written).toContain(
			'<cx:f>Sheet1!$F$2:$F$4</cx:f><cx:lvl ptCount="3" formatCode="0.0"><cx:pt idx="0">10</cx:pt><cx:pt idx="2">30</cx:pt></cx:lvl>',
		);
		expect(written).toContain(
			'<cx:f>Sheet1!$E$2:$E$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt><cx:pt idx="2">C</cx:pt></cx:lvl>',
		);
		expect(() => setChartExSeriesValues(space, 0, [Number.NaN])).toThrow(RangeError);
		expect(() => setChartExSeriesValues(space, 3, [1])).toThrow(RangeError);
	});

	it('gives a series without data a new data set, with the value dimension of its layout', () => {
		const treemap = structuredClone(space);
		treemap.plotArea.series[0]!.layoutId = 'treemap';
		delete treemap.plotArea.series[0]!.dataId;
		const edited = setChartExCategories(setChartExSeriesValues(treemap, 0, [1, 2]), [
			['a', 'b'],
			['P', 'P'],
		]);
		expect(edited.plotArea.series[0]!.dataId).toBe(1);
		const points = (values: string[]) => values.map((value, index) => ({ index, value }));
		expect(edited.chartData.data[1]).toEqual({
			id: 1,
			dimensions: [
				{
					kind: 'numDim',
					type: 'size',
					levels: [{ pointCount: 2, formatCode: 'General', points: points(['1', '2']) }],
				},
				{
					kind: 'strDim',
					type: 'cat',
					levels: [
						{ pointCount: 2, points: points(['a', 'b']) },
						{ pointCount: 2, points: points(['P', 'P']) },
					],
				},
			],
		});
	});

	it('sets, retexts and removes the title', () => {
		expect(write(setChartExTitle(space, 'Q4 & more'))).toContain(
			'<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr /><a:p><a:r><a:t>Q4 &amp; more</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>',
		);
		expect(write(setChartExTitle(space, undefined))).not.toContain('<cx:title');
		const rich = parseChartExSpace(FULL).chartExSpace;
		const retexted = setChartExTitle(rich, 'New');
		expect(retexted.title?.tx?.rich?.text).toBe('New');
		expect(retexted.title?.spPr).toEqual(rich.title?.spPr);
		expect(writeChartExSpace(retexted)).toContain(
			'<cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>New</a:t></a:r></a:p></cx:rich>',
		);
	});

	it('moves, hides and adds the legend', () => {
		expect(write(setChartExLegend(space, { visible: true, position: 'tr' }))).toContain(
			'<cx:legend pos="r" align="ctr" overlay="0" />',
		);
		const hidden = setChartExLegend(space, { visible: false });
		expect(hidden.legend).toBeUndefined();
		expect(setChartExLegend(hidden, { visible: true }).legend).toEqual({
			position: 'r',
			align: 'ctr',
			overlay: false,
		});
	});

	it('toggles data labels on every series', () => {
		const shown = setChartExDataLabelsVisible(space, true);
		expect(write(shown)).toContain(
			'<cx:dataLabels><cx:visibility seriesName="0" categoryName="1" value="1" /></cx:dataLabels><cx:dataId val="0" />',
		);
		expect(
			setChartExDataLabelsVisible(shown, false).plotArea.series[0]!.dataLabels,
		).toBeUndefined();
	});

	it('edits the series fill', () => {
		const fill: DrawingFill = {
			kind: 'solid',
			color: { kind: 'srgb', value: 'FF0000', transforms: [] },
		};
		expect(write(setChartExSeriesFill(space, 0, fill))).toContain(
			'<cx:spPr><a:solidFill><a:srgbClr val="FF0000" /></a:solidFill></cx:spPr><cx:dataId',
		);
	});

	it('edits the family options', () => {
		const waterfall = setChartExWaterfall(space, 0, {
			subtotals: [2, 1, 2],
			connectorLines: false,
		});
		expect(write(waterfall)).toContain(
			'<cx:layoutPr><cx:visibility connectorLines="0" /><cx:subtotals><cx:idx val="1" /><cx:idx val="2" /></cx:subtotals></cx:layoutPr>',
		);
		expect(
			setChartExWaterfall(space, 0, { subtotals: [] }).plotArea.series[0]!.layoutProperties,
		).toEqual({ visibility: { connectorLines: true } });
		const binned = setChartExBinning(space, 0, { binSize: 5, binCount: 3, intervalClosed: 'l' });
		expect(binned.plotArea.series[0]!.layoutProperties?.binning).toEqual({
			binSize: 5,
			intervalClosed: 'l',
		});
		const aggregated = setChartExBinning(binned, 0, 'category').plotArea.series[0]!;
		expect(aggregated.layoutProperties?.aggregation).toBe(true);
		expect(aggregated.layoutProperties?.binning).toBeUndefined();
		expect(
			setChartExParentLabelLayout(space, 0, 'banner').plotArea.series[0]!.layoutProperties
				?.parentLabelLayout,
		).toBe('banner');
		expect(
			setChartExBoxWhisker(space, 0, { quartileMethod: 'inclusive', outliers: false }).plotArea
				.series[0]!.layoutProperties,
		).toMatchObject({
			visibility: { connectorLines: true, outliers: false },
			statistics: { quartileMethod: 'inclusive' },
		});
	});
});
