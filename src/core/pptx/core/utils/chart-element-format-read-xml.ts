/**
 * Reads a {@link PptxChartElementFormat} back from a ChartML `c:spPr` (the
 * inverse of `chart-element-format-xml.ts`), and every chart element's format
 * from a chart part, so a reopened chart shows its transparency, line
 * details, effects and legend / axis / data-label boxes in the Format pane
 * and on the slide.
 *
 * Only what the panes model is read. A gradient is read when the caller
 * passes a gradient parser; a picture fill records its relationship id, and
 * the load path resolves it to an image.
 *
 * @module utils/chart-element-format-read-xml
 */
import type {
	PptxChartGradientFill,
	PptxChartData,
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementFormats,
	PptxChartElementLine,
	PptxChartLineEndType,
	PptxChartTextFill,
	XmlObject,
} from '../types';
import { readChartThreeD } from './chart-element-format-3d-xml';
import { findChartAxisNode } from './chart-element-format-targets';
import { readChartTextFormat } from './chart-element-format-text-xml';

type LocalName = (key: string) => string;
/** Resolves a fill wrapper (`a:solidFill`) to the hex it paints. */
export type ResolveFillColor = (fill: XmlObject) => string | undefined;
/** Parses the gradient of an `a:gradFill` node's parent (`c:spPr`, `a:ln`, `a:rPr`). */
export type ParseChartElementGradient = (parent: XmlObject) => PptxChartGradientFill | undefined;

/** How colours and gradients are resolved while reading. */
export interface ChartElementFormatReaders {
	resolve: ResolveFillColor;
	parseGradient?: ParseChartElementGradient;
}

const EMU_PER_POINT = 12700;

function child(node: XmlObject | undefined, name: string, localName: LocalName) {
	if (!node) {
		return undefined;
	}
	const key = Object.keys(node).find((k) => !k.startsWith('@_') && localName(k) === name);
	if (!key) {
		return undefined;
	}
	const value = node[key];
	// `<a:noFill/>` parses to an empty string: presence is what counts.
	return value && typeof value === 'object' && !Array.isArray(value) ? (value as XmlObject) : {};
}

function has(node: XmlObject | undefined, name: string, localName: LocalName): boolean {
	return Boolean(
		node && Object.keys(node).some((k) => !k.startsWith('@_') && localName(k) === name),
	);
}

const round = (value: number) => Math.round(value * 100) / 100;
const points = (emu: unknown): number | undefined => {
	const n = Number(emu);
	return Number.isFinite(n) ? round(n / EMU_PER_POINT) : undefined;
};

/** Transparency (0..100) of the colour choice inside a fill wrapper. */
function transparencyOf(fill: XmlObject, localName: LocalName): number | undefined {
	for (const [key, value] of Object.entries(fill)) {
		if (key.startsWith('@_') || !value || typeof value !== 'object') {
			continue;
		}
		const alpha = child(value as XmlObject, 'alpha', localName)?.['@_val'];
		const n = Number(alpha);
		if (alpha !== undefined && Number.isFinite(n)) {
			const transparency = round(100 - n / 1000);
			return transparency > 0 ? transparency : undefined;
		}
	}
	return undefined;
}

function paint(
	node: XmlObject,
	localName: LocalName,
	readers: ChartElementFormatReaders,
): PptxChartElementFill | undefined {
	if (has(node, 'noFill', localName)) {
		return { kind: 'none' };
	}
	const solid = child(node, 'solidFill', localName);
	if (solid) {
		const color = readers.resolve(solid);
		const transparency = transparencyOf(solid, localName);
		return {
			kind: 'solid',
			...(color ? { color } : {}),
			...(transparency !== undefined ? { transparency } : {}),
		};
	}
	if (has(node, 'gradFill', localName)) {
		const gradient = readers.parseGradient?.(node);
		return gradient ? { kind: 'gradient', gradient } : undefined;
	}
	const pattern = child(node, 'pattFill', localName);
	if (pattern) {
		const fg = child(pattern, 'fgClr', localName);
		const bg = child(pattern, 'bgClr', localName);
		return {
			kind: 'pattern',
			pattern: {
				preset: String(pattern['@_prst'] ?? 'pct5'),
				foreground: (fg && readers.resolve(fg)) ?? '#000000',
				background: (bg && readers.resolve(bg)) ?? '#FFFFFF',
			},
		};
	}
	const blipFill = child(node, 'blipFill', localName);
	if (blipFill) {
		const blip = child(blipFill, 'blip', localName);
		const relId = blip
			? Object.entries(blip).find(
					([key]) => key.startsWith('@_') && localName(key.slice(2)) === 'embed',
				)?.[1]
			: undefined;
		const amt = Number(child(blip, 'alphaModFix', localName)?.['@_amt']);
		return {
			kind: 'picture',
			picture: {
				...(typeof relId === 'string' && relId ? { relId } : {}),
				...(has(blipFill, 'tile', localName) ? { tile: true } : {}),
				...(Number.isFinite(amt) && amt < 100000 ? { transparency: round(100 - amt / 1000) } : {}),
			},
		};
	}
	return undefined;
}

const LINE_ENDS = new Set(['triangle', 'stealth', 'diamond', 'oval', 'arrow', 'none']);

function readLine(
	ln: XmlObject,
	localName: LocalName,
	readers: ChartElementFormatReaders,
): PptxChartElementLine {
	const fill = paint(ln, localName, readers);
	const line: PptxChartElementLine =
		fill && (fill.kind === 'none' || fill.kind === 'solid' || fill.kind === 'gradient')
			? (fill as PptxChartElementLine)
			: { kind: 'automatic' };
	const width = points(ln['@_w']);
	if (ln['@_w'] !== undefined && width !== undefined) line.width = width;
	const cmpd = ln['@_cmpd'];
	if (
		cmpd === 'sng' ||
		cmpd === 'dbl' ||
		cmpd === 'thickThin' ||
		cmpd === 'thinThick' ||
		cmpd === 'tri'
	) {
		line.compound = cmpd;
	}
	const cap = ln['@_cap'];
	if (cap === 'rnd' || cap === 'sq' || cap === 'flat') line.cap = cap;
	const dash = child(ln, 'prstDash', localName)?.['@_val'];
	if (typeof dash === 'string') line.dash = dash;
	for (const join of ['round', 'bevel', 'miter'] as const) {
		if (has(ln, join, localName)) line.join = join;
	}
	const head = child(ln, 'headEnd', localName)?.['@_type'];
	if (typeof head === 'string' && LINE_ENDS.has(head) && head !== 'none') {
		line.beginArrow = head as PptxChartLineEndType;
	}
	const tail = child(ln, 'tailEnd', localName)?.['@_type'];
	if (typeof tail === 'string' && LINE_ENDS.has(tail) && tail !== 'none') {
		line.endArrow = tail as PptxChartLineEndType;
	}
	return line;
}

function readEffects(
	list: XmlObject,
	localName: LocalName,
	resolve: ResolveFillColor,
): PptxChartElementEffects | undefined {
	const effects: PptxChartElementEffects = {};
	const shadow = child(list, 'outerShdw', localName);
	if (shadow) {
		const color = resolve(shadow) ?? '#000000';
		const transparency = transparencyOf(shadow, localName);
		effects.shadow = {
			color,
			...(transparency !== undefined ? { transparency } : {}),
			blur: points(shadow['@_blurRad'] ?? 0) ?? 0,
			distance: points(shadow['@_dist'] ?? 0) ?? 0,
			angle: round((Number(shadow['@_dir'] ?? 0) || 0) / 60000),
		};
	}
	const glow = child(list, 'glow', localName);
	if (glow) {
		const transparency = transparencyOf(glow, localName);
		effects.glow = {
			color: resolve(glow) ?? '#FFFFFF',
			...(transparency !== undefined ? { transparency } : {}),
			radius: points(glow['@_rad'] ?? 0) ?? 0,
		};
	}
	const soft = child(list, 'softEdge', localName);
	if (soft) {
		effects.softEdge = points(soft['@_rad'] ?? 0) ?? 0;
	}
	return Object.keys(effects).length > 0 ? effects : undefined;
}

/** The format a `c:spPr` describes, or `undefined` when it carries nothing the panes model. */
export function readChartElementFormatFromSpPr(
	spPr: XmlObject | undefined,
	localName: LocalName,
	readers: ChartElementFormatReaders | ResolveFillColor,
): PptxChartElementFormat | undefined {
	if (!spPr) {
		return undefined;
	}
	const r = typeof readers === 'function' ? { resolve: readers } : readers;
	const format: PptxChartElementFormat = {};
	const fill = paint(spPr, localName, r);
	if (fill) format.fill = fill;
	const ln = child(spPr, 'ln', localName);
	if (ln) format.line = readLine(ln, localName, r);
	const effectLst = child(spPr, 'effectLst', localName);
	const effects = effectLst ? readEffects(effectLst, localName, r.resolve) : undefined;
	if (effects) format.effects = effects;
	const threeD = readChartThreeD(spPr, localName, r.resolve);
	if (threeD) format.threeD = threeD;
	return Object.keys(format).length > 0 ? format : undefined;
}

/** A text fill as Text Options reads it (pattern and picture text fills are not offered). */
function textPaint(
	node: XmlObject,
	localName: LocalName,
	readers: ChartElementFormatReaders,
): PptxChartTextFill | undefined {
	const fill = paint(node, localName, readers);
	return fill && fill.kind !== 'pattern' && fill.kind !== 'picture'
		? (fill as PptxChartTextFill)
		: undefined;
}

/** The format of an element owner: its `c:spPr` plus, for a text owner, its Text Options. */
function readOwnerFormat(
	owner: XmlObject | undefined,
	localName: LocalName,
	readers: ChartElementFormatReaders,
	withText: boolean,
): PptxChartElementFormat | undefined {
	const format = readChartElementFormatFromSpPr(
		child(owner, 'spPr', localName),
		localName,
		readers,
	);
	const text = withText
		? readChartTextFormat(
				owner,
				localName,
				(node) => textPaint(node, localName, readers),
				(list) => readEffects(list, localName, readers.resolve),
			)
		: undefined;
	return text ? { ...format, text } : format;
}

/** The nodes of a chart part a format belongs to. */
export interface ChartElementFormatNodes {
	chartSpace: XmlObject | undefined;
	/** `c:ser` nodes, index-aligned with `PptxChartData.series`. */
	seriesNodes: readonly XmlObject[];
}

/** Every `c:ser` of a plot area in document order (chart groups in order). */
export function collectChartSeriesNodes(
	plotArea: XmlObject | undefined,
	localName: LocalName,
): XmlObject[] {
	const out: XmlObject[] = [];
	for (const [key, value] of Object.entries(plotArea ?? {})) {
		if (!localName(key).endsWith('Chart') || !value || typeof value !== 'object') {
			continue;
		}
		// Two groups of one type (primary and secondary axis) parse as an array.
		for (const group of Array.isArray(value) ? value : [value]) {
			if (!group || typeof group !== 'object') continue;
			for (const [childKey, ser] of Object.entries(group as XmlObject)) {
				if (localName(childKey) !== 'ser') continue;
				for (const node of Array.isArray(ser) ? ser : [ser]) {
					if (node && typeof node === 'object') out.push(node as XmlObject);
				}
			}
		}
	}
	return out;
}

/**
 * Every chart element's format, keyed like `PptxChartData.elementFormats`.
 * `data` supplies the axes the axis keys index. Returns `undefined` when no
 * element carries a format the panes model.
 */
export function readChartElementFormatsFromXml(
	nodes: ChartElementFormatNodes,
	data: Pick<PptxChartData, 'axes'>,
	localName: LocalName,
	readers: ChartElementFormatReaders | ResolveFillColor,
): PptxChartElementFormats | undefined {
	const out: PptxChartElementFormats = {};
	const r = typeof readers === 'function' ? { resolve: readers } : readers;
	const spPrOf = (owner: XmlObject | undefined) => readOwnerFormat(owner, localName, r, false);
	const textOf = (owner: XmlObject | undefined) => readOwnerFormat(owner, localName, r, true);
	const put = (key: string, format: PptxChartElementFormat | undefined) => {
		if (format) out[key] = format;
	};
	const chartRoot = child(nodes.chartSpace, 'chart', localName);
	const plotArea = child(chartRoot, 'plotArea', localName);
	put('chartArea', textOf(nodes.chartSpace));
	put('plotArea', spPrOf(plotArea));
	put('title', textOf(child(chartRoot, 'title', localName)));
	put('legend', textOf(child(chartRoot, 'legend', localName)));
	(data.axes ?? []).forEach((_, axisIndex) => {
		const axis = plotArea ? findChartAxisNode(plotArea, data, axisIndex, localName) : undefined;
		put(`axis:${axisIndex}`, textOf(axis));
		put(`majorGridlines:${axisIndex}`, spPrOf(child(axis, 'majorGridlines', localName)));
		put(`minorGridlines:${axisIndex}`, spPrOf(child(axis, 'minorGridlines', localName)));
	});
	nodes.seriesNodes.forEach((ser, index) => put(`series:${index}`, spPrOf(ser)));
	// Chart-wide data labels: the first chart group's own `c:dLbls`, else the first series'.
	for (const [key, value] of Object.entries(plotArea ?? {})) {
		if (localName(key).endsWith('Chart') && value && typeof value === 'object') {
			const group = (Array.isArray(value) ? value[0] : value) as XmlObject | undefined;
			const own = textOf(child(group, 'dLbls', localName));
			if (own) {
				put('dataLabels', own);
				break;
			}
		}
	}
	if (!out.dataLabels && nodes.seriesNodes[0]) {
		put('dataLabels', textOf(child(nodes.seriesNodes[0], 'dLbls', localName)));
	}
	return Object.keys(out).length > 0 ? out : undefined;
}
