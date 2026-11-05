/**
 * Fields of a NEW chart the legacy generator dropped, written through the shared chart core: each
 * case builds the chart with the public SDK (or sets the model field the SDK documents), saves,
 * checks the XML, reloads with `new PptxHandler().load` and reads the field back. Omitted, zero,
 * fractional and boundary values are cases of their own.
 */
import { describe, expect, it } from 'vitest';

import { setChartAxis, setChartAxisTitleStyle } from '../../core/builders/sdk/chart-operations';
import { withChartElementFormat } from '../../core/builders/sdk/chart-element-format-operations';
import type { ChartPptxElement, PptxChartData, PptxChartType } from '../../core/types';
import { firstElement, newChartDeck } from '../helpers/new-chart-deck';

interface FieldCase {
	name: string;
	type?: PptxChartType;
	edit: (element: ChartPptxElement, data: PptxChartData) => void;
	/** Fragments the saved part contains. */
	xml?: string[];
	/** Fragments the saved part must not contain. */
	absent?: string[];
	read: (chart: PptxChartData, xml: string) => unknown;
	want: unknown;
}

const valueAxis = (chart: PptxChartData) => chart.axes?.find((axis) => axis.axisType === 'valAx');

const CASES: FieldCase[] = [
	// Axis titles, number formats and title fonts.
	{
		name: 'axis title and number format',
		edit: (element) => {
			setChartAxis(element, 'valAx', { title: 'Revenue (k)', numberFormat: '0.0%' });
			setChartAxis(element, 'catAx', { title: 'Quarter' });
		},
		xml: ['<c:numFmt formatCode="0.0%" sourceLinked="0"></c:numFmt>', '<a:t>Revenue (k)</a:t>'],
		read: (chart) => [
			valueAxis(chart)?.titleText,
			valueAxis(chart)?.numFmt,
			chart.axes?.find((axis) => axis.axisType === 'catAx')?.titleText,
		],
		want: ['Revenue (k)', { formatCode: '0.0%', sourceLinked: false }, 'Quarter'],
	},
	{
		name: 'axis without title or number format',
		edit: (element) => setChartAxis(element, 'valAx', { min: 0 }),
		absent: ['<c:numFmt', '<c:title>'],
		read: (chart) => [valueAxis(chart)?.titleText, valueAxis(chart)?.numFmt, valueAxis(chart)?.min],
		want: [undefined, undefined, 0],
	},
	{
		name: 'an empty axis title is no title',
		edit: (element) => setChartAxis(element, 'valAx', { title: '' }),
		absent: ['<c:title>'],
		read: (chart) => valueAxis(chart)?.titleText,
		want: undefined,
	},
	{
		name: 'axis title font',
		edit: (element) => {
			setChartAxis(element, 'valAx', { title: 'Rev' });
			setChartAxisTitleStyle(element, 'valAx', {
				fontSize: 14,
				fontBold: true,
				fontColor: '#FF0000',
			});
		},
		read: (chart) => [
			valueAxis(chart)?.fontSize,
			valueAxis(chart)?.fontBold,
			valueAxis(chart)?.fontColor,
		],
		want: [14, true, '#FF0000'],
	},
	{
		name: 'fractional axis title size',
		edit: (element) => {
			setChartAxis(element, 'valAx', { title: 'Rev' });
			setChartAxisTitleStyle(element, 'valAx', { fontSize: 10.5 });
		},
		xml: ['sz="1050"'],
		read: (chart) => valueAxis(chart)?.fontSize,
		want: 10.5,
	},
	// Chart title.
	{
		name: 'title runs win over the flat title',
		edit: (_, data) => {
			data.title = 'ignored';
			data.titleRuns = [
				{ text: 'Q4 ', bold: true },
				{ text: 'sales', italic: true, fontSize: 20, color: '#C00000' },
			];
		},
		absent: ['ignored'],
		read: (chart) =>
			chart.titleRuns?.map((run) => [run.text, run.bold, run.italic, run.fontSize, run.color]),
		want: [
			['Q4 ', true, undefined, undefined, undefined],
			['sales', undefined, true, 20, '#C00000'],
		],
	},
	{
		name: 'hasTitle false removes the title',
		edit: (_, data) => {
			data.title = 'Hidden';
			data.style = { ...data.style, hasTitle: false };
		},
		xml: ['<c:autoTitleDeleted val="1"></c:autoTitleDeleted>'],
		absent: ['<c:title>', 'Hidden'],
		read: (chart) => chart.title,
		want: undefined,
	},
	{
		name: 'an empty title is no title',
		edit: (_, data) => {
			data.title = '';
		},
		absent: ['<c:title>'],
		read: (chart) => chart.title,
		want: undefined,
	},
	{
		name: 'title font',
		edit: (_, data) => {
			data.title = 'Sales';
			Object.assign(data.style!, {
				titleFontSize: 20,
				titleFontBold: true,
				titleFontColor: '#00AA00',
			});
		},
		read: (chart) => [
			chart.style?.titleFontSize,
			chart.style?.titleFontBold,
			chart.style?.titleFontColor,
		],
		want: [20, true, '#00AA00'],
	},
	// Legend and chart text.
	{
		name: 'legend overlay and text style',
		edit: (_, data) => {
			Object.assign(data.style!, {
				legendOverlay: true,
				legendTextStyle: { fontSize: 9, bold: true, color: '#404040' },
			});
		},
		xml: ['<c:overlay val="1"></c:overlay>'],
		read: (chart) => [chart.style?.legendOverlay, chart.style?.legendTextStyle],
		want: [true, expect.objectContaining({ fontSize: 9, bold: true, color: '#404040' })],
	},
	{
		name: 'hidden legend entry',
		edit: (_, data) => {
			data.style!.legendEntries = [{ index: 1, deleted: true }];
		},
		read: (chart) => chart.style?.legendEntries,
		want: [expect.objectContaining({ index: 1, deleted: true })],
	},
	{
		name: 'no legend',
		edit: (_, data) => {
			data.style!.hasLegend = false;
		},
		absent: ['<c:legend>'],
		read: (chart) => chart.style?.hasLegend === true,
		want: false,
	},
	{
		name: 'chart text style',
		edit: (_, data) => {
			data.style!.textStyle = { fontSize: 11, color: '#333333' };
		},
		read: (chart) => [chart.style?.textStyle?.fontSize, chart.style?.textStyle?.color],
		want: [11, '#333333'],
	},
	// Element formats.
	{
		name: 'axis line format, with the model listing the value axis first',
		edit: (element) => {
			setChartAxis(element, 'valAx', { min: 0 });
			setChartAxis(element, 'catAx', { title: 'Q' });
			element.chartData = withChartElementFormat(
				element.chartData!,
				{ kind: 'axis', axisIndex: 0 },
				{ line: { kind: 'solid', color: '#FF0000', width: 2 } },
			);
		},
		read: (_, xml) => [
			firstElement(xml, 'c:valAx').includes('FF0000'),
			firstElement(xml, 'c:catAx').includes('FF0000'),
		],
		want: [true, false],
	},
	{
		name: 'gridline and axis formats on a chart with no axis entries name its axes in order',
		edit: (element) => {
			element.chartData = withChartElementFormat(
				element.chartData!,
				{ kind: 'axis', axisIndex: 1 },
				{ line: { kind: 'solid', color: '#0000FF', width: 1 } },
			);
		},
		read: (_, xml) => [
			firstElement(xml, 'c:valAx').includes('0000FF'),
			firstElement(xml, 'c:catAx').includes('0000FF'),
		],
		want: [true, false],
	},
	{
		name: 'hidden axis and major gridlines',
		edit: (element) => {
			setChartAxis(element, 'catAx', { visible: false });
			setChartAxis(element, 'valAx', { majorGridlines: true });
		},
		xml: ['<c:delete val="1"></c:delete>', '<c:majorGridlines>'],
		read: (chart) => [
			chart.axes?.find((axis) => axis.axisType === 'catAx')?.deleted,
			valueAxis(chart)?.majorGridlines,
		],
		want: [true, true],
	},
	{
		name: 'legend at the bottom, the title text escaped',
		edit: (_, data) => {
			data.title = 'Sales & <costs> "Q4"';
			data.style!.legendPosition = 'b';
		},
		xml: ['<c:legendPos val="b"></c:legendPos>', '<a:t>Sales &amp; &lt;costs&gt;'],
		read: (chart) => [chart.title, chart.style?.legendPosition],
		want: ['Sales & <costs> "Q4"', 'b'],
	},
	{
		name: 'trendline, error bars, marker and a point fill',
		type: 'line',
		edit: (_, data) => {
			Object.assign(data.series[0]!, {
				trendlines: [{ trendlineType: 'linear', displayEq: true }],
				errBars: [{ direction: 'y', barType: 'both', valType: 'fixedVal', val: 1.5 }],
				marker: { symbol: 'diamond', size: 9 },
			});
			data.series[1]!.dataPoints = [{ idx: 2, spPr: { fillColor: '#00B050' } }];
		},
		read: (chart) => [
			chart.series[0]!.trendlines?.[0]?.trendlineType,
			chart.series[0]!.trendlines?.[0]?.displayEq,
			chart.series[0]!.errBars?.[0]?.val,
			chart.series[0]!.marker?.symbol,
			chart.series[0]!.marker?.size,
			chart.series[1]!.dataPoints?.find((point) => point.idx === 2)?.spPr?.fillColor,
		],
		want: ['linear', true, 1.5, 'diamond', 9, '#00B050'],
	},
	{
		name: 'data table, up-down bars and protection',
		type: 'line',
		edit: (_, data) => {
			data.dataTable = { showHorzBorder: true, showKeys: false };
			data.upDownBars = { gapWidth: 80 };
			data.protection = { chartObject: true };
		},
		xml: ['<c:dTable>', '<c:upDownBars><c:gapWidth val="80"></c:gapWidth>', '<c:protection>'],
		read: (chart) => [
			chart.dataTable?.showHorzBorder,
			chart.dataTable?.showKeys,
			chart.upDownBars?.gapWidth,
			chart.protection?.chartObject,
		],
		want: [true, false, 80, true],
	},
	// Chrome and blanks.
	{
		name: 'blanks drawn as zero, left out of the cache',
		type: 'line',
		edit: (_, data) => {
			data.chartChrome = { dispBlanksAs: 'zero' };
			data.series[0]!.blanks = [false, true, false];
		},
		xml: [
			'<c:dispBlanksAs val="zero"></c:dispBlanksAs>',
			'<c:ptCount val="3"></c:ptCount><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2">',
		],
		read: (chart) => [chart.chartChrome?.dispBlanksAs, chart.series[0]!.blanks],
		want: ['zero', [false, true, false]],
	},
	{
		name: 'blanks spanned, labels over the maximum, hidden cells plotted',
		type: 'line',
		edit: (_, data) => {
			data.chartChrome = { dispBlanksAs: 'span', showDLblsOverMax: true };
			data.plotVisibleOnly = false;
		},
		read: (chart) => [
			chart.chartChrome?.dispBlanksAs,
			chart.chartChrome?.showDLblsOverMax,
			chart.plotVisibleOnly,
		],
		want: ['span', true, false],
	},
	// Series lines.
	{
		name: 'series line width, dash and smoothing',
		type: 'line',
		edit: (_, data) =>
			Object.assign(data.series[0]!, {
				color: '#C00000',
				lineWidth: 3,
				lineDashStyle: 'dash',
				smooth: true,
			}),
		xml: [
			'<a:ln w="38100">',
			'<a:prstDash val="dash"></a:prstDash>',
			'<c:smooth val="1"></c:smooth>',
		],
		read: (chart) => [
			chart.series[0]!.color,
			chart.series[0]!.lineWidth,
			chart.series[0]!.lineDashStyle,
			chart.series[0]!.smooth,
		],
		want: ['#C00000', 3, 'dash', true],
	},
	{
		name: 'fractional line width',
		type: 'line',
		edit: (_, data) => {
			data.series[1]!.lineWidth = 0.25;
		},
		xml: ['<a:ln w="3175">'],
		read: (chart) => chart.series[1]!.lineWidth,
		want: 0.25,
	},
	{
		name: 'zero line width, written as the thinnest line and read back unset',
		type: 'line',
		edit: (_, data) => {
			data.series[0]!.lineWidth = 0;
		},
		xml: ['<a:ln w="0">'],
		read: (chart) => chart.series[0]!.lineWidth,
		want: undefined,
	},
	{
		name: 'smoothing switched off',
		type: 'line',
		edit: (_, data) => {
			data.series[0]!.smooth = false;
		},
		xml: ['<c:smooth val="0"></c:smooth>'],
		read: (chart) => chart.series[0]!.smooth,
		want: false,
	},
	// Scatter and bubble data.
	{
		name: 'scatter x values',
		type: 'scatter',
		edit: (_, data) => {
			data.series[0]!.xValues = [1.5, 2.5, -7];
		},
		read: (chart) => [chart.series[0]!.xValues, chart.series[1]!.xValues],
		want: [
			[1.5, 2.5, -7],
			[1, 2, 3],
		],
	},
	{
		name: 'bubble sizes, zero and fractional',
		type: 'bubble',
		edit: (_, data) => {
			data.series[0]!.bubbleSizes = [0, 0.5, 2];
		},
		read: (chart) => [chart.series[0]!.bubbleSizes, chart.series[1]!.bubbleSizes],
		want: [
			[0, 0.5, 2],
			[1, 1, 1],
		],
	},
	// 3-D.
	{
		name: '3-D view',
		type: 'bar3D',
		edit: (_, data) => {
			data.view3D = { rotX: 20, rotY: 40, rAngAx: false, perspective: 45, depthPercent: 150 };
		},
		xml: ['<c:view3D><c:rotX val="20"></c:rotX>'],
		read: (chart) => chart.view3D,
		want: expect.objectContaining({
			rotX: 20,
			rotY: 40,
			rAngAx: false,
			perspective: 45,
			depthPercent: 150,
		}),
	},
	{
		name: 'a 3-D chart without a view writes none',
		type: 'bar3D',
		edit: () => {},
		absent: ['<c:view3D>'],
		read: (chart) => chart.view3D,
		want: undefined,
	},
	// Group settings the generator ignored.
	{
		name: 'pie of pie options',
		type: 'ofPie',
		edit: (_, data) => {
			data.ofPieOptions = {
				ofPieType: 'bar',
				splitType: 'pos',
				splitPos: 2,
				secondPieSize: 75,
				serLines: true,
			};
		},
		xml: ['<c:ofPieType val="bar"></c:ofPieType>', '<c:serLines>'],
		read: (chart) => chart.ofPieOptions,
		want: expect.objectContaining({
			ofPieType: 'bar',
			splitType: 'pos',
			splitPos: 2,
			secondPieSize: 75,
			serLines: true,
		}),
	},
	{
		name: 'pie slice explosion, and zero',
		type: 'pie',
		edit: (_, data) => {
			data.series[0]!.explosion = 25;
			data.series[1]!.explosion = 0;
		},
		read: (chart) => [chart.series[0]!.explosion, chart.series[1]!.explosion ?? 0],
		want: [25, 0],
	},
	{
		name: 'scatter style, radar style and colour variation',
		type: 'scatter',
		edit: (_, data) => {
			data.scatterStyle = 'smoothMarker';
			data.varyColors = true;
		},
		xml: [
			'<c:scatterStyle val="smoothMarker"></c:scatterStyle>',
			'<c:varyColors val="1"></c:varyColors>',
		],
		read: (chart) => [chart.scatterStyle, chart.varyColors],
		want: ['smoothMarker', true],
	},
	{
		name: 'filled radar',
		type: 'radar',
		edit: (_, data) => {
			data.radarStyle = 'filled';
		},
		read: (chart) => chart.radarStyle,
		want: 'filled',
	},
	{
		name: 'surface drawn filled',
		type: 'surface',
		edit: (_, data) => {
			data.wireframe = false;
		},
		xml: ['<c:wireframe val="0"></c:wireframe>'],
		read: (chart) => chart.wireframe,
		want: false,
	},
	{
		name: '3-D bar shape and gap depth',
		type: 'bar3D',
		edit: (_, data) => {
			data.barShape = 'cylinder';
			data.gapDepth = 200;
		},
		xml: ['<c:gapDepth val="200"></c:gapDepth><c:shape val="cylinder"></c:shape>'],
		read: (chart) => [chart.barShape, chart.gapDepth],
		want: ['cylinder', 200],
	},
	{
		name: 'bubble options, the scale rounded to a whole percent',
		type: 'bubble',
		edit: (_, data) => {
			data.bubbleOptions = { bubbleScale: 150.4, showNegativeBubbles: false, sizeRepresents: 'w' };
		},
		xml: ['<c:bubbleScale val="150"></c:bubbleScale>'],
		read: (chart) => chart.bubbleOptions,
		want: expect.objectContaining({
			bubbleScale: 150,
			showNegativeBubbles: false,
			sizeRepresents: 'w',
		}),
	},
	{
		name: 'stock high-low lines removed',
		type: 'stock',
		edit: (_, data) => {
			data.hiLowLines = null;
		},
		absent: ['<c:hiLowLines'],
		read: (chart) => chart.hiLowLines ?? null,
		want: null,
	},
	// Dates.
	{
		name: 'date categories on a date axis',
		type: 'line',
		edit: (element, data) => {
			data.dateCategories = { values: [45000, 45031, 45061], formatCode: 'mmm yy' };
			setChartAxis(element, 'dateAx', { title: 'Month' });
		},
		xml: ['<c:dateAx>', '<c:formatCode>mmm yy</c:formatCode>'],
		read: (chart) => [
			chart.dateCategories?.values,
			chart.dateCategories?.formatCode,
			chart.axes?.[0]?.axisType,
			chart.axes?.[0]?.titleText,
		],
		want: [[45000, 45031, 45061], 'mmm yy', 'dateAx', 'Month'],
	},
	{
		name: 'date categories without a date axis stay on a category axis',
		type: 'line',
		edit: (_, data) => {
			data.dateCategories = { values: [45000, 45031, 45061], formatCode: 'mmm yy' };
		},
		xml: ['<c:catAx>', '<c:formatCode>mmm yy</c:formatCode>'],
		read: (chart) => chart.axes?.[0]?.axisType,
		want: 'catAx',
	},
	{
		name: '1904 date system',
		edit: (_, data) => {
			data.date1904 = true;
		},
		xml: ['<c:date1904 val="1"></c:date1904>'],
		read: (chart) => chart.date1904,
		want: true,
	},
];

describe('new chart fields written through the shared chart core', () => {
	it.each(CASES)('$name', async (item) => {
		const { element, save } = await newChartDeck(item.type ?? 'bar');
		item.edit(element, element.chartData!);
		const { xml, chart, warnings } = await save();
		expect(warnings).toStrictEqual([]);
		for (const fragment of item.xml ?? []) expect(xml).toContain(fragment);
		for (const fragment of item.absent ?? []) expect(xml).not.toContain(fragment);
		expect(item.read(chart, xml)).toStrictEqual(item.want);
	});
});
