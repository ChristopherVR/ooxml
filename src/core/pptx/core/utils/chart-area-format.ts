/**
 * Fill and border of a chart container's `c:spPr`: the chart area
 * (`c:chartSpace/c:spPr`) or the plot area (`c:plotArea/c:spPr`).
 *
 * The model keeps a solid fill or `'none'` (`PptxChartStyle.chartAreaFill` /
 * `plotAreaFill`), a gradient that wins over it (`chartAreaGradient` /
 * `plotAreaGradient`) and the border colour or `'none'` (`chartAreaBorder` /
 * `plotAreaBorder`). This module parses that shape, builds it for a generated
 * chart, and reconciles it into a loaded chart's `c:spPr` touching only the
 * part (fill or border) the model changed, so an unedited chart keeps its
 * authored XML byte-for-byte.
 *
 * @module utils/chart-area-format
 */
import type { PptxChartGradientFill, XmlObject } from '../types';
import type { ResolveChartColor } from './chart-color-choice';
import { chartColorChoiceValue, chartColorHex } from './chart-color-choice';
import type { ParseChartGradient } from './chart-gradient-fill-writer';
import { applyChartGradientToSpPr, buildChartGradFillXml } from './chart-gradient-fill-writer';
import { chartGradientsEqual } from './chart-gradient-fill-xml';
import { colorsEqual } from './color-xml-preservation';

type GetLocalName = (key: string) => string;
type XmlValue = XmlObject[string];

/** A chart or plot area's fill and border, as modelled. */
export interface ChartAreaFormat {
	/** Solid fill colour, or `'none'` for `a:noFill`. */
	fill?: string;
	/** Gradient fill; wins over {@link fill}. */
	gradient?: PptxChartGradientFill;
	/** Border colour (`a:ln/a:solidFill`), or `'none'` for `a:ln/a:noFill`. */
	border?: string;
}

const FILL_CHOICE = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
const AFTER_LN = new Set(['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']);
const LN_FILL_CHOICE = new Set(['noFill', 'solidFill', 'gradFill', 'pattFill']);

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((key) => getLocalName(key) === local);
}

function asObject(value: unknown): XmlObject | undefined {
	return value && typeof value === 'object' && !Array.isArray(value)
		? (value as XmlObject)
		: undefined;
}

/** `'none'` for a direct `a:noFill`, the colour of a direct `a:solidFill`, else `undefined`. */
function parsePaint(
	node: XmlObject | undefined,
	getLocalName: GetLocalName,
	resolveColor: ResolveChartColor,
): string | undefined {
	if (!node) {
		return undefined;
	}
	// `<a:noFill/>` parses to an empty string, so presence decides, not truthiness.
	if (findKey(node, 'noFill', getLocalName) !== undefined) {
		return 'none';
	}
	const solid = asObject(node[findKey(node, 'solidFill', getLocalName) ?? '']);
	return solid ? resolveColor(solid) : undefined;
}

/** Parse the fill, gradient and border of `owner/c:spPr`. */
export function parseChartAreaFormat(
	owner: XmlObject | undefined,
	getLocalName: GetLocalName,
	resolveColor: ResolveChartColor,
	parseGradient?: ParseChartGradient,
): ChartAreaFormat {
	const spPr = owner ? asObject(owner[findKey(owner, 'spPr', getLocalName) ?? '']) : undefined;
	if (!spPr) {
		return {};
	}
	const format: ChartAreaFormat = {};
	const fill = parsePaint(spPr, getLocalName, resolveColor);
	if (fill !== undefined) {
		format.fill = fill;
	}
	const gradient = parseGradient?.(spPr);
	if (gradient) {
		format.gradient = gradient;
	}
	const ln = asObject(spPr[findKey(spPr, 'ln', getLocalName) ?? '']);
	const border = parsePaint(ln, getLocalName, resolveColor);
	if (border !== undefined) {
		format.border = border;
	}
	return format;
}

function paintXml(paint: string): XmlObject {
	return paint === 'none'
		? { 'a:noFill': {} }
		: { 'a:solidFill': { 'a:srgbClr': { '@_val': chartColorHex(paint) } } };
}

/** `c:spPr` content for a generated chart, or `undefined` when the model says nothing. */
export function buildChartAreaSpPr(format: ChartAreaFormat): XmlObject | undefined {
	const spPr: XmlObject = {};
	if (format.gradient) {
		spPr['a:gradFill'] = buildChartGradFillXml(format.gradient);
	} else if (format.fill !== undefined) {
		Object.assign(spPr, paintXml(format.fill));
	}
	if (format.border !== undefined) {
		spPr['a:ln'] = paintXml(format.border);
	}
	return Object.keys(spPr).length > 0 ? spPr : undefined;
}

function samePaint(a: string | undefined, b: string | undefined): boolean {
	if (a === undefined || b === undefined || a === 'none' || b === 'none') {
		return a === b;
	}
	return colorsEqual(a, b);
}

function sameFill(model: ChartAreaFormat, authored: ChartAreaFormat): boolean {
	if (model.gradient || authored.gradient) {
		return Boolean(
			model.gradient && authored.gradient && chartGradientsEqual(model.gradient, authored.gradient),
		);
	}
	return samePaint(model.fill, authored.fill);
}

/** Replace `node`'s fill choice (members of `choices`) with `entry`, before the `after` keys. */
function setPaint(
	node: XmlObject,
	entry: [string, XmlValue] | undefined,
	choices: Set<string>,
	after: Set<string>,
	getLocalName: GetLocalName,
): void {
	const entries = Object.entries(node).filter(([key]) => !choices.has(getLocalName(key)));
	if (entry) {
		const at = entries.findIndex(([key]) => after.has(getLocalName(key)));
		entries.splice(at === -1 ? entries.length : at, 0, entry);
	}
	for (const key of Object.keys(node)) {
		delete node[key];
	}
	for (const [key, value] of entries) {
		node[key] = value;
	}
}

function paintEntry(
	paint: string | undefined,
	authored: XmlObject,
	getLocalName: GetLocalName,
	resolveColor: ResolveChartColor,
): [string, XmlValue] | undefined {
	if (paint === undefined) {
		return undefined;
	}
	if (paint === 'none') {
		return ['a:noFill', {}];
	}
	const solid = asObject(authored[findKey(authored, 'solidFill', getLocalName) ?? '']);
	return ['a:solidFill', chartColorChoiceValue(solid, paint, resolveColor)];
}

/** Options for {@link applyChartAreaFormatToXml}. */
export interface ChartAreaFormatWriteOptions {
	resolveColor: ResolveChartColor;
	parseGradient?: ParseChartGradient;
	/** Attach a newly created `c:spPr` to `owner` at its schema position. */
	insertSpPr: (owner: XmlObject, spPr: XmlObject) => void;
}

/**
 * Reconcile `owner/c:spPr` with the model. Returns whether anything changed;
 * an authored node the model still describes is left untouched.
 */
export function applyChartAreaFormatToXml(
	owner: XmlObject,
	format: ChartAreaFormat,
	getLocalName: GetLocalName,
	options: ChartAreaFormatWriteOptions,
): boolean {
	const authored = parseChartAreaFormat(
		owner,
		getLocalName,
		options.resolveColor,
		options.parseGradient,
	);
	const fillChanged = !sameFill(format, authored);
	const borderChanged = !samePaint(format.border, authored.border);
	if (!fillChanged && !borderChanged) {
		return false;
	}
	let spPr = asObject(owner[findKey(owner, 'spPr', getLocalName) ?? '']);
	if (!spPr) {
		spPr = {};
		options.insertSpPr(owner, spPr);
	}
	if (fillChanged) {
		if (format.gradient) {
			applyChartGradientToSpPr(spPr, format.gradient, getLocalName, options);
		} else {
			const entry = paintEntry(format.fill, spPr, getLocalName, options.resolveColor);
			setPaint(spPr, entry, FILL_CHOICE, new Set(['ln', ...AFTER_LN]), getLocalName);
		}
	}
	if (borderChanged) {
		const lnKey = findKey(spPr, 'ln', getLocalName);
		const ln = asObject(lnKey ? spPr[lnKey] : undefined) ?? {};
		const entry = paintEntry(format.border, ln, getLocalName, options.resolveColor);
		setPaint(
			ln,
			entry,
			LN_FILL_CHOICE,
			new Set(['prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd', 'extLst']),
			getLocalName,
		);
		if (Object.keys(ln).length === 0) {
			if (lnKey) {
				delete spPr[lnKey];
			}
		} else if (!lnKey) {
			setPaint(spPr, ['a:ln', ln], new Set(), AFTER_LN, getLocalName);
		}
	}
	if (Object.keys(spPr).length === 0) {
		const spPrKey = findKey(owner, 'spPr', getLocalName);
		if (spPrKey) {
			delete owner[spPrKey];
		}
	}
	return true;
}
