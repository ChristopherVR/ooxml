/**
 * Reads a {@link PptxChartElementFormat} back from a ChartML `c:spPr` (the
 * inverse of `chart-element-format-xml.ts`), and every chart element's format
 * from a chart part, so a reopened chart shows its transparency, line
 * details, effects and legend / axis / data-label boxes in the Format pane
 * and on the slide.
 *
 * Only what the panes model is read: a gradient, picture or pattern fill
 * leaves `fill` unset (the chart's gradient fields carry gradients).
 *
 * @module utils/chart-element-format-read-xml
 */
import type {
	PptxChartData,
	PptxChartElementEffects,
	PptxChartElementFill,
	PptxChartElementFormat,
	PptxChartElementFormats,
	PptxChartElementLine,
	PptxChartLineEndType,
	XmlObject,
} from '../types';
import { findChartAxisNode } from './chart-element-format-targets';

type LocalName = (key: string) => string;
/** Resolves a fill wrapper (`a:solidFill`) to the hex it paints. */
export type ResolveFillColor = (fill: XmlObject) => string | undefined;

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
	resolve: ResolveFillColor,
): { kind: 'none' | 'solid'; color?: string; transparency?: number } | undefined {
	if (has(node, 'noFill', localName)) {
		return { kind: 'none' };
	}
	const solid = child(node, 'solidFill', localName);
	if (!solid) {
		return undefined;
	}
	const color = resolve(solid);
	const transparency = transparencyOf(solid, localName);
	return {
		kind: 'solid',
		...(color ? { color } : {}),
		...(transparency !== undefined ? { transparency } : {}),
	};
}

const LINE_ENDS = new Set(['triangle', 'stealth', 'diamond', 'oval', 'arrow', 'none']);

function readLine(
	ln: XmlObject,
	localName: LocalName,
	resolve: ResolveFillColor,
): PptxChartElementLine {
	const line: PptxChartElementLine = paint(ln, localName, resolve) ?? { kind: 'automatic' };
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
	resolve: ResolveFillColor,
): PptxChartElementFormat | undefined {
	if (!spPr) {
		return undefined;
	}
	const format: PptxChartElementFormat = {};
	const fill = paint(spPr, localName, resolve);
	if (fill) format.fill = fill as PptxChartElementFill;
	const ln = child(spPr, 'ln', localName);
	if (ln) format.line = readLine(ln, localName, resolve);
	const effectLst = child(spPr, 'effectLst', localName);
	const effects = effectLst ? readEffects(effectLst, localName, resolve) : undefined;
	if (effects) format.effects = effects;
	return Object.keys(format).length > 0 ? format : undefined;
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
		for (const [childKey, ser] of Object.entries(value as XmlObject)) {
			if (localName(childKey) !== 'ser') continue;
			for (const node of Array.isArray(ser) ? ser : [ser]) {
				if (node && typeof node === 'object') out.push(node as XmlObject);
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
	resolve: ResolveFillColor,
): PptxChartElementFormats | undefined {
	const out: PptxChartElementFormats = {};
	const spPrOf = (owner: XmlObject | undefined) =>
		readChartElementFormatFromSpPr(child(owner, 'spPr', localName), localName, resolve);
	const put = (key: string, format: PptxChartElementFormat | undefined) => {
		if (format) out[key] = format;
	};
	const chartRoot = child(nodes.chartSpace, 'chart', localName);
	const plotArea = child(chartRoot, 'plotArea', localName);
	put('chartArea', spPrOf(nodes.chartSpace));
	put('plotArea', spPrOf(plotArea));
	put('title', spPrOf(child(chartRoot, 'title', localName)));
	put('legend', spPrOf(child(chartRoot, 'legend', localName)));
	(data.axes ?? []).forEach((_, axisIndex) => {
		const axis = plotArea ? findChartAxisNode(plotArea, data, axisIndex, localName) : undefined;
		put(`axis:${axisIndex}`, spPrOf(axis));
		put(`majorGridlines:${axisIndex}`, spPrOf(child(axis, 'majorGridlines', localName)));
		put(`minorGridlines:${axisIndex}`, spPrOf(child(axis, 'minorGridlines', localName)));
	});
	nodes.seriesNodes.forEach((ser, index) => put(`series:${index}`, spPrOf(ser)));
	// Chart-wide data labels: the first chart group's own `c:dLbls`, else the first series'.
	for (const [key, value] of Object.entries(plotArea ?? {})) {
		if (localName(key).endsWith('Chart') && value && typeof value === 'object') {
			const own = spPrOf(child(value as XmlObject, 'dLbls', localName));
			if (own) {
				put('dataLabels', own);
				break;
			}
		}
	}
	if (!out.dataLabels && nodes.seriesNodes[0]) {
		put('dataLabels', spPrOf(child(nodes.seriesNodes[0], 'dLbls', localName)));
	}
	return Object.keys(out).length > 0 ? out : undefined;
}
