// The chart style and colour style parts a new extended chart carries. Word and PowerPoint refuse
// to open a `chartEx` part without both relationships (observed with Office 16: the document fails
// to open when either is missing), unlike a classic chart part. The colour style is Office's
// colourful palette 10 (`color-palettes`); the chart style is a schema-complete
// `cs:chartStyle` ([MS-ODRAWXML] 2.6, chartStyle namespace) whose entries hold Office's documented
// default look: automatic series colours (`cs:styleClr val="auto"`), light grey axis and gridlines
// and grey text on the minor font. Its id is 0 (none of Office's numbered styles), so the
// product's Chart Styles gallery shows no style as current.
import { findChartColorPalette } from '../color-palettes';
import { chartColorStyleXml } from '../color-style';

export const CHART_STYLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
export const CHART_STYLE_CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml';

const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const text = (lum: [number, number] | undefined) =>
	lum
		? `<a:schemeClr val="tx1"><a:lumMod val="${lum[0]}"/><a:lumOff val="${lum[1]}"/></a:schemeClr>`
		: '<a:schemeClr val="tx1"/>';
const grey = (lumMod: number, lumOff: number) =>
	`<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="${lumMod}"/><a:lumOff val="${lumOff}"/></a:schemeClr></a:solidFill>`;
const line = (width: number, fill: string) =>
	`<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:round/></a:ln>`;

interface Entry {
	fill?: 'auto';
	line?: 'auto';
	font?: [number, number];
	spPr?: string;
	size?: number;
	bold?: boolean;
}

function entry(name: string, value: Entry = {}): string {
	const ln = value.line
		? '<cs:lnRef idx="0"><cs:styleClr val="auto"/></cs:lnRef>'
		: '<cs:lnRef idx="0"/>';
	const fill = value.fill
		? '<cs:fillRef idx="0"><cs:styleClr val="auto"/></cs:fillRef>'
		: '<cs:fillRef idx="0"/>';
	const rpr =
		value.size !== undefined
			? `<cs:defRPr sz="${value.size}"${value.bold === undefined ? '' : ` b="${value.bold ? 1 : 0}"`}/>`
			: '';
	return `<cs:${name}>${ln}${fill}<cs:effectRef idx="0"/><cs:fontRef idx="minor">${text(value.font)}</cs:fontRef>${value.spPr ? `<cs:spPr>${value.spPr}</cs:spPr>` : ''}${rpr}</cs:${name}>`;
}

const GREY_TEXT: [number, number] = [65000, 35000];
const AXIS_LINE = line(9525, grey(15000, 85000));
const SERIES_FILL = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';

/** The chart style part of a new extended chart. */
export function chartExStyleXml(): string {
	const entries = [
		entry('axisTitle', { font: GREY_TEXT, size: 900 }),
		entry('categoryAxis', { font: GREY_TEXT, spPr: AXIS_LINE, size: 900 }),
		entry('chartArea', {
			spPr: `<a:solidFill><a:schemeClr val="bg1"/></a:solidFill>${AXIS_LINE}`,
			size: 1000,
		}),
		entry('dataLabel', { font: GREY_TEXT, size: 900 }),
		entry('dataPoint', { fill: 'auto', spPr: SERIES_FILL }),
		entry('dataPoint3D', { fill: 'auto', spPr: SERIES_FILL }),
		entry('dataPointLine', {
			line: 'auto',
			spPr: `<a:ln w="28575" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln>`,
		}),
		entry('dataPointMarker', { fill: 'auto', spPr: SERIES_FILL }),
		entry('dataPointWireframe', {
			line: 'auto',
			spPr: `<a:ln w="9525" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:round/></a:ln>`,
		}),
		entry('dataTable', { font: GREY_TEXT, spPr: AXIS_LINE, size: 900 }),
		entry('downBar', { spPr: `<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>` }),
		entry('dropLine', { spPr: line(9525, grey(35000, 65000)) }),
		entry('errorBar', { spPr: line(9525, grey(65000, 35000)) }),
		entry('floor', {}),
		entry('gridlineMajor', { spPr: AXIS_LINE }),
		entry('gridlineMinor', { spPr: line(9525, grey(5000, 95000)) }),
		entry('hiLoLine', { spPr: line(9525, grey(75000, 25000)) }),
		entry('leaderLine', { spPr: line(9525, grey(35000, 65000)) }),
		entry('legend', { font: GREY_TEXT, size: 900 }),
		entry('plotArea', {}),
		entry('plotArea3D', {}),
		entry('seriesAxis', { font: GREY_TEXT, size: 900 }),
		entry('seriesLine', { spPr: line(9525, grey(35000, 65000)) }),
		entry('title', { font: GREY_TEXT, size: 1400 }),
		entry('trendline', {
			line: 'auto',
			spPr: `<a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`,
		}),
		entry('trendlineLabel', { font: GREY_TEXT, size: 900 }),
		entry('upBar', { spPr: `<a:solidFill><a:schemeClr val="lt1"/></a:solidFill>` }),
		entry('valueAxis', { font: GREY_TEXT, size: 900 }),
		entry('wall', {}),
	];
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cs:chartStyle xmlns:cs="${CS}" xmlns:a="${A}" id="0">${entries.join('')}</cs:chartStyle>`;
}

/** The colour style part of a new extended chart: Office's colourful palette (id 10). */
export function chartExColorStyleXml(): string {
	const palette = findChartColorPalette(10);
	if (!palette) throw new Error('The default chart colour palette is missing.');
	return chartColorStyleXml(palette);
}
