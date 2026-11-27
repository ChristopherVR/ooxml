/**
 * Native charts for the samples. A chart is described by a small spec, written as a DrawingML
 * chart part and read into the core's neutral `ChartSpace` model, which the session inserts and
 * the workbook writer serializes. Building the model this way (instead of the summary fields of
 * `addChart`) is what allows combination charts, secondary axes and explicit fonts.
 */
import type { ChartSpace } from '../../../../src/core/chart/model';
import { parseChartSpace } from '../../../../src/core/chart/parse-space';
import { parseAddress, parseRange } from '../../../../src/core/xlsx/address';
import { getValue } from '../../../../src/core/xlsx/cells';
import type { ChartObject, DrawingAnchor, Workbook } from '../../../../src/core/xlsx/model';
import { sheetByName } from '../../../../src/core/xlsx/workbook';
import { FONTS } from '../../lib/fiction';
import { xmlEscape } from '../../lib/package';

export interface SeriesSpec {
	/** The cell holding the series name (`'Chart data'!$C$4`). */
	name: string;
	/** Category and value ranges, sheet-qualified and absolute. */
	categories: string;
	values: string;
	color: string;
	/** Per-point colours (pie and doughnut slices, highlighted bars). */
	pointColors?: string[];
	/** Line series: width in points (default 2.25) and whether to draw round markers. */
	lineWidth?: number;
	markers?: boolean;
	smooth?: boolean;
	/** Area series: fill opacity in percent. */
	opacity?: number;
	/** Show value labels with this number format. */
	labels?: string;
}

export interface GroupSpec {
	type: 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut';
	grouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	series: SeriesSpec[];
	gapWidth?: number;
	overlap?: number;
	/** Plot on the secondary value axis (right edge). */
	secondary?: boolean;
	holeSize?: number;
}

export interface ChartSpec {
	title?: string;
	groups: GroupSpec[];
	legend?: 'b' | 'r' | 't' | 'l';
	/** Number formats of the primary and secondary value axes. */
	valueFormat?: string;
	secondaryFormat?: string;
	categoryFormat?: string;
	/** Colours of the text, the gridlines and the chart area (`RRGGBB`). */
	ink: string;
	grid: string;
	background?: string;
	/** Axis bounds, when the automatic ones waste space. */
	valueMin?: number;
	valueMax?: number;
	secondaryMin?: number;
	secondaryMax?: number;
}

const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const solid = (hex: string, opacity?: number) =>
	`<a:solidFill><a:srgbClr val="${hex}">${opacity === undefined ? '' : `<a:alpha val="${opacity * 1000}"/>`}</a:srgbClr></a:solidFill>`;
const face = (name: string) =>
	`<a:latin typeface="${name}"/><a:ea typeface="${name}"/><a:cs typeface="${name}"/>`;
const runProps = (tag: string, size: number, color: string, font: string, bold = false) =>
	`<a:${tag} lang="en-US" sz="${size * 100}" b="${bold ? 1 : 0}">${solid(color)}${face(font)}</a:${tag}>`;
const textProps = (size: number, color: string, font: string = FONTS.body) =>
	`<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${runProps('defRPr', size, color, font)}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
const line = (hex: string, width: number) =>
	`<a:ln w="${Math.round(width * 12700)}" cap="rnd">${solid(hex)}<a:round/></a:ln>`;
const noLine = '<a:ln><a:noFill/></a:ln>';

/** The stored values of a sheet-qualified range, row by row. */
function cached(workbook: Workbook, ref: string): (string | number | null)[] {
	const bang = ref.lastIndexOf('!');
	const sheet = sheetByName(workbook, ref.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"));
	const text = ref.slice(bang + 1).replace(/\$/g, '');
	const range = parseRange(text) ?? (parseAddress(text) && parseRange(`${text}:${text}`));
	if (!sheet || !range) throw new Error(`Chart reference ${ref} does not resolve`);
	const values: (string | number | null)[] = [];
	for (let row = range.start.row; row <= range.end.row; row++)
		for (let col = range.start.col; col <= range.end.col; col++) {
			const value = getValue(sheet, row, col);
			values.push(typeof value === 'number' || typeof value === 'string' ? value : null);
		}
	return values;
}

function source(workbook: Workbook, ref: string, format?: string): string {
	const values = cached(workbook, ref);
	const points = values
		.map((value, i) =>
			value === null ? '' : `<c:pt idx="${i}"><c:v>${xmlEscape(String(value))}</c:v></c:pt>`,
		)
		.join('');
	const numeric = values.every((value) => value === null || typeof value === 'number');
	return numeric
		? `<c:numRef><c:f>${xmlEscape(ref)}</c:f><c:numCache><c:formatCode>${xmlEscape(format ?? 'General')}</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`
		: `<c:strRef><c:f>${xmlEscape(ref)}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

function seriesXml(
	workbook: Workbook,
	group: GroupSpec,
	series: SeriesSpec,
	index: number,
	spec: ChartSpec,
): string {
	const linear = group.type === 'line';
	const round = group.type === 'pie' || group.type === 'doughnut';
	const gap = spec.background ?? 'FFFFFF';
	const name = cached(workbook, series.name)[0] ?? '';
	const shape = linear
		? `<c:spPr>${line(series.color, series.lineWidth ?? 2.25)}</c:spPr>`
		: `<c:spPr>${solid(series.color, series.opacity)}${round ? line(gap, 1.5) : noLine}</c:spPr>`;
	const marker = !linear
		? ''
		: series.markers
			? `<c:marker><c:symbol val="circle"/><c:size val="6"/><c:spPr>${solid(gap)}${line(series.color, 1.5)}</c:spPr></c:marker>`
			: '<c:marker><c:symbol val="none"/></c:marker>';
	const points = (series.pointColors ?? [])
		.map(
			(color, i) =>
				`<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${solid(color)}${round ? line(gap, 1.5) : noLine}</c:spPr></c:dPt>`,
		)
		.join('');
	const labels = series.labels
		? `<c:dLbls><c:numFmt formatCode="${xmlEscape(series.labels)}" sourceLinked="0"/><c:spPr><a:noFill/>${noLine}</c:spPr>${textProps(8, round ? 'FFFFFF' : spec.ink)}<c:showLegendKey val="0"/><c:showVal val="${round ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${round ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`
		: '';
	return (
		`<c:ser><c:idx val="${index}"/><c:order val="${index}"/>` +
		`<c:tx><c:strRef><c:f>${xmlEscape(series.name)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${xmlEscape(String(name))}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
		shape +
		(group.type === 'column' || group.type === 'bar' ? '<c:invertIfNegative val="0"/>' : '') +
		marker +
		points +
		labels +
		`<c:cat>${source(workbook, series.categories, spec.categoryFormat)}</c:cat>` +
		`<c:val>${source(workbook, series.values)}</c:val>` +
		(linear ? `<c:smooth val="${series.smooth ? 1 : 0}"/>` : '') +
		'</c:ser>'
	);
}

const AXES = { category: 501, value: 502, secondaryCategory: 503, secondaryValue: 504 };

function groupXml(workbook: Workbook, group: GroupSpec, first: number, spec: ChartSpec): string {
	const series = group.series
		.map((entry, i) => seriesXml(workbook, group, entry, first + i, spec))
		.join('');
	const axes = group.secondary
		? `<c:axId val="${AXES.secondaryCategory}"/><c:axId val="${AXES.secondaryValue}"/>`
		: `<c:axId val="${AXES.category}"/><c:axId val="${AXES.value}"/>`;
	switch (group.type) {
		case 'column':
		case 'bar': {
			const grouping = group.grouping ?? 'clustered';
			const overlap = group.overlap ?? (grouping === 'clustered' ? undefined : 100);
			return `<c:barChart><c:barDir val="${group.type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${series}<c:gapWidth val="${group.gapWidth ?? 60}"/>${overlap === undefined ? '' : `<c:overlap val="${overlap}"/>`}${axes}</c:barChart>`;
		}
		case 'line':
			return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${series}<c:marker val="1"/>${axes}</c:lineChart>`;
		case 'area':
			return `<c:areaChart><c:grouping val="${group.grouping ?? 'standard'}"/><c:varyColors val="0"/>${series}${axes}</c:areaChart>`;
		case 'pie':
			return `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
		case 'doughnut':
			return `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="${group.holeSize ?? 62}"/></c:doughnutChart>`;
	}
}

function axesXml(spec: ChartSpec): string {
	const horizontal = spec.groups.some((group) => group.type === 'bar');
	const text = textProps(9, spec.ink);
	const axisLine = `<c:spPr><a:noFill/>${line(spec.grid, 0.75)}</c:spPr>`;
	const scaling = (min?: number, max?: number, reversed = false) =>
		`<c:scaling><c:orientation val="${reversed ? 'maxMin' : 'minMax'}"/>${max === undefined ? '' : `<c:max val="${max}"/>`}${min === undefined ? '' : `<c:min val="${min}"/>`}</c:scaling>`;
	// A bar chart lists its first category at the top, as a ranked table reads.
	const category = (id: number, cross: number, hidden: boolean) =>
		`<c:catAx><c:axId val="${id}"/>${scaling(undefined, undefined, horizontal)}<c:delete val="${hidden ? 1 : 0}"/><c:axPos val="${horizontal ? 'l' : 'b'}"/>` +
		`<c:numFmt formatCode="${xmlEscape(spec.categoryFormat ?? 'General')}" sourceLinked="${spec.categoryFormat ? 0 : 1}"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/>${axisLine}${text}` +
		`<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
	const value = (id: number, cross: number, secondary: boolean) =>
		`<c:valAx><c:axId val="${id}"/>${secondary ? scaling(spec.secondaryMin, spec.secondaryMax) : scaling(spec.valueMin, spec.valueMax)}<c:delete val="0"/><c:axPos val="${horizontal ? 'b' : secondary ? 'r' : 'l'}"/>` +
		(secondary
			? ''
			: `<c:majorGridlines><c:spPr>${line(spec.grid, 0.5)}</c:spPr></c:majorGridlines>`) +
		`<c:numFmt formatCode="${xmlEscape((secondary ? spec.secondaryFormat : spec.valueFormat) ?? 'General')}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/>${noLine}</c:spPr>${text}` +
		`<c:crossAx val="${cross}"/><c:crosses val="${secondary || horizontal ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
	let xml = category(AXES.category, AXES.value, false) + value(AXES.value, AXES.category, false);
	if (spec.groups.some((group) => group.secondary))
		xml +=
			value(AXES.secondaryValue, AXES.secondaryCategory, true) +
			category(AXES.secondaryCategory, AXES.secondaryValue, true);
	return xml;
}

/** The neutral model of the chart `spec` describes, with caches read from the workbook. */
export function chartSpace(workbook: Workbook, spec: ChartSpec): ChartSpace {
	let first = 0;
	const groups = spec.groups
		.map((group) => {
			const xml = groupXml(workbook, group, first, spec);
			first += group.series.length;
			return xml;
		})
		.join('');
	const round = spec.groups.every((group) => group.type === 'pie' || group.type === 'doughnut');
	const title = spec.title
		? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${runProps('defRPr', 11, spec.ink, FONTS.display, true)}</a:pPr><a:r>${runProps('rPr', 11, spec.ink, FONTS.display, true)}<a:t>${xmlEscape(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`
		: '';
	const legend = spec.legend
		? `<c:legend><c:legendPos val="${spec.legend}"/><c:overlay val="0"/>${textProps(9, spec.ink)}</c:legend>`
		: '';
	const area = spec.background
		? `<c:spPr>${solid(spec.background)}${noLine}</c:spPr>`
		: `<c:spPr><a:noFill/>${noLine}</c:spPr>`;
	const xml =
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
		`<c:date1904 val="0"/><c:roundedCorners val="0"/><c:chart>${title}<c:autoTitleDeleted val="${spec.title ? 0 : 1}"/>` +
		`<c:plotArea><c:layout/>${groups}${round ? '' : axesXml(spec)}<c:spPr><a:noFill/>${noLine}</c:spPr></c:plotArea>` +
		`${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${area}${textProps(9, spec.ink)}</c:chartSpace>`;
	const parsed = parseChartSpace(xml);
	if (parsed.issues.length)
		throw new Error(
			`Chart model issues: ${parsed.issues.map((issue) => issue.message).join('; ')}`,
		);
	return parsed.chartSpace;
}

/** A two-cell anchor from the top-left of cell `from` to the top-left of cell `to`. */
export function anchor(from: string, to: string): DrawingAnchor {
	const a = parseAddress(from);
	const b = parseAddress(to);
	if (!a || !b) throw new Error(`Bad anchor ${from}:${to}`);
	return {
		from: { row: a.row, col: a.col, rowOffset: 0, colOffset: 0 },
		to: { row: b.row, col: b.col, rowOffset: 0, colOffset: 0 },
	};
}

/** What `EditSession.addChart` takes for a chart built from a spec. */
export function chartObject(
	workbook: Workbook,
	name: string,
	description: string,
	at: DrawingAnchor,
	spec: ChartSpec,
): Omit<ChartObject, 'kind'> {
	return {
		anchor: at,
		name,
		description,
		// The summary is derived from the model when the chart is inserted.
		chartType: 'column',
		series: [],
		showLegend: Boolean(spec.legend),
		chartSpace: chartSpace(workbook, spec),
	};
}
