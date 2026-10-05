import { NS, buildXml, children, elements, first, parseXml } from '../../xml/index.js';
import type { ChartObject, ChartSeries, Color } from '../model.js';
import { XML_HEADER, escapeAttr, escapeText } from './xml-out.js';

const THEME_NAMES = [
	'bg1',
	'tx1',
	'bg2',
	'tx2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
];

function colorFill(color: Color | undefined, index: number): string {
	if (color?.rgb) return `<a:solidFill><a:srgbClr val="${color.rgb.slice(-6)}"/></a:solidFill>`;
	const name = THEME_NAMES[color?.theme ?? 4 + (index % 6)] ?? 'accent1';
	return `<a:solidFill><a:schemeClr val="${name}"/></a:solidFill>`;
}

const pt = (values: readonly (string | number | null)[]) =>
	values
		.map((value, idx) =>
			value === null || value === ''
				? ''
				: `<c:pt idx="${idx}"><c:v>${escapeText(String(value))}</c:v></c:pt>`,
		)
		.join('');

function strSource(ref: string | undefined, values: readonly (string | number)[]): string {
	const cache = `<c:ptCount val="${values.length}"/>${pt(values)}`;
	return ref
		? `<c:strRef><c:f>${escapeText(ref)}</c:f><c:strCache>${cache}</c:strCache></c:strRef>`
		: `<c:strLit>${cache}</c:strLit>`;
}

function numSource(ref: string | undefined, values: readonly (string | number | null)[]): string {
	const cache = `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${pt(values)}`;
	return ref
		? `<c:numRef><c:f>${escapeText(ref)}</c:f><c:numCache>${cache}</c:numCache></c:numRef>`
		: `<c:numLit>${cache}</c:numLit>`;
}

function seriesXml(chart: ChartObject, series: ChartSeries, index: number): string {
	const type = chart.chartType;
	let out = `<c:idx val="${index}"/><c:order val="${index}"/>`;
	if (series.nameRef)
		out += `<c:tx>${strSource(series.nameRef, series.name === undefined ? [] : [series.name])}</c:tx>`;
	else if (series.name !== undefined) out += `<c:tx><c:v>${escapeText(series.name)}</c:v></c:tx>`;
	const lineLike = type === 'line' || type === 'scatter' || type === 'radar';
	const fill = colorFill(series.color, index);
	if (type !== 'pie' && type !== 'doughnut')
		out += lineLike
			? `<c:spPr><a:ln w="28575" cap="rnd">${fill}</a:ln></c:spPr>`
			: `<c:spPr>${fill}</c:spPr>`;
	if (lineLike) out += '<c:marker><c:symbol val="none"/></c:marker>';
	if (type === 'bar' || type === 'column') out += '<c:invertIfNegative val="0"/>';
	const numericCats =
		series.categories.length > 0 && series.categories.every((c) => typeof c === 'number');
	const hasCats = series.categoriesRef !== undefined || series.categories.length > 0;
	const cats = numericCats
		? numSource(series.categoriesRef, series.categories)
		: strSource(series.categoriesRef, series.categories);
	if (type === 'scatter') {
		if (hasCats) out += `<c:xVal>${cats}</c:xVal>`;
		out += `<c:yVal>${numSource(series.valuesRef, series.values)}</c:yVal><c:smooth val="0"/>`;
		return `<c:ser>${out}</c:ser>`;
	}
	if (hasCats) out += `<c:cat>${cats}</c:cat>`;
	out += `<c:val>${numSource(series.valuesRef, series.values)}</c:val>`;
	if (type === 'line') out += '<c:smooth val="0"/>';
	return `<c:ser>${out}</c:ser>`;
}

const axes = (horizontal: boolean, scatter: boolean) => {
	const common = '<c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>';
	const ticks =
		'<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>';
	const first = scatter
		? `<c:valAx><c:axId val="500000001"/>${common}<c:axPos val="b"/>${ticks}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
		: `<c:catAx><c:axId val="500000001"/>${common}<c:axPos val="${horizontal ? 'l' : 'b'}"/>${ticks}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
	const second = `<c:valAx><c:axId val="500000002"/>${common}<c:axPos val="${horizontal ? 'b' : 'l'}"/><c:majorGridlines/>${ticks}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="${scatter ? 'midCat' : 'between'}"/></c:valAx>`;
	return first + second;
};
const AX_IDS = '<c:axId val="500000001"/><c:axId val="500000002"/>';

function plotXml(chart: ChartObject): string {
	const series = chart.series.map((s, i) => seriesXml(chart, s, i)).join('');
	const grouping =
		chart.grouping ??
		(chart.chartType === 'bar' || chart.chartType === 'column' ? 'clustered' : 'standard');
	switch (chart.chartType) {
		case 'line':
			return `<c:lineChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${AX_IDS}</c:lineChart>${axes(false, false)}`;
		case 'area':
			return `<c:areaChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${series}${AX_IDS}</c:areaChart>${axes(false, false)}`;
		case 'pie':
			return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
		case 'doughnut':
			return `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
		case 'scatter':
			return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${AX_IDS}</c:scatterChart>${axes(false, true)}`;
		case 'radar':
			return `<c:radarChart><c:radarStyle val="marker"/><c:varyColors val="0"/>${series}${AX_IDS}</c:radarChart>${axes(false, false)}`;
		default: {
			const horizontal = chart.chartType === 'bar';
			const bar = grouping === 'standard' ? 'clustered' : grouping;
			const overlap = bar === 'clustered' ? '' : '<c:overlap val="100"/>';
			return `<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${bar}"/><c:varyColors val="0"/>${series}<c:gapWidth val="150"/>${overlap}${AX_IDS}</c:barChart>${axes(horizontal, false)}`;
		}
	}
}

/** A new chart part for a chart created in the model (no source part to keep). */
export function chartXml(chart: ChartObject): string {
	const title =
		chart.title !== undefined
			? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeText(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
			: '<c:autoTitleDeleted val="1"/>';
	const legend = chart.showLegend
		? `<c:legend><c:legendPos val="${escapeAttr(chart.legendPosition ?? 'r')}"/><c:overlay val="0"/></c:legend>`
		: '';
	return `${XML_HEADER}<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plotXml(chart)}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

/**
 * Updates the series references of a kept chart part when the model moved them (rows inserted,
 * sheet renamed). Returns `undefined` when nothing changed, so the part stays byte-identical.
 */
export function patchChartReferences(xml: string, chart: ChartObject): string | undefined {
	const doc = parseXml(xml, { label: 'XLSX chart' });
	const plotArea = first(first(doc.documentElement, 'chart', NS.c), 'plotArea', NS.c);
	const plot = plotArea
		? elements(plotArea).find((node) => children(node, 'ser', NS.c).length)
		: undefined;
	if (!plot) return undefined;
	let changed = false;
	const setRef = (node: Element | undefined, ref: string | undefined) => {
		const f = node
			? first(first(node, 'numRef', NS.c) ?? first(node, 'strRef', NS.c), 'f', NS.c)
			: undefined;
		if (f && ref !== undefined && f.textContent !== ref) {
			f.textContent = ref;
			changed = true;
		}
	};
	children(plot, 'ser', NS.c).forEach((ser, index) => {
		const model = chart.series[index];
		if (!model) return;
		setRef(first(ser, 'tx', NS.c), model.nameRef);
		setRef(first(ser, 'cat', NS.c) ?? first(ser, 'xVal', NS.c), model.categoriesRef);
		setRef(first(ser, 'val', NS.c) ?? first(ser, 'yVal', NS.c), model.valuesRef);
	});
	return changed ? buildXml(doc) : undefined;
}
