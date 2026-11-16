import { parseDrawingPictureFill } from './picture-fill';
// DrawingML fill and line readers over the DOM.
import { parseDrawingColorIn } from './drawing-color';
import { NS, children, elements, first, type XmlElement } from './dom';
import { readDrawingGradient } from './gradient-read';
import { DOM_DRAWING_VIEW } from './node-view';
import type { DrawingFill, DrawingLine } from './types';
import { buildXml } from '../xml/index';

const FILL_ELEMENTS = new Set([
	'noFill',
	'solidFill',
	'gradFill',
	'pattFill',
	'blipFill',
	'grpFill',
]);

/** The fill element held directly by a shape-properties or line element, if any. */
export function fillElementOf(properties: XmlElement | undefined): XmlElement | undefined {
	return properties
		? elements(properties).find(
				(element) =>
					FILL_ELEMENTS.has(element.localName) &&
					(!element.namespaceURI || element.namespaceURI === NS.a),
			)
		: undefined;
}

/** Reads the fill of `a:spPr`/`a:ln`; `undefined` when the element declares none (inherit). */
export function parseDrawingFill(properties: XmlElement | undefined): DrawingFill | undefined {
	const fill = fillElementOf(properties);
	if (!fill) return undefined;
	switch (fill.localName) {
		case 'noFill':
			return { kind: 'none' };
		case 'solidFill': {
			const color = parseDrawingColorIn(fill);
			return color ? { kind: 'solid', color } : { kind: 'unsupported', element: 'solidFill' };
		}
		case 'gradFill': {
			const { kind, ...gradient } = readDrawingGradient(DOM_DRAWING_VIEW, fill).fill;
			return { kind, sourceXml: buildXml(fill), ...gradient };
		}
		case 'pattFill': {
			const foreground = parseDrawingColorIn(first(fill, 'fgClr', NS.a));
			const background = parseDrawingColorIn(first(fill, 'bgClr', NS.a));
			return {
				kind: 'pattern',
				sourceXml: buildXml(fill),
				preset: fill.getAttribute('prst') ?? '',
				...(foreground ? { foreground } : {}),
				...(background ? { background } : {}),
			};
		}
		case 'blipFill':
			return parseDrawingPictureFill(fill);
		default:
			return { kind: 'unsupported', element: fill.localName };
	}
}

/** Reads `a:ln`: width in EMU, fill, dash and cap. */
export function parseDrawingLine(ln: XmlElement | undefined): DrawingLine | undefined {
	if (!ln) return undefined;
	const line: DrawingLine = { sourceXml: buildXml(ln) };
	const width = Number.parseInt(ln.getAttribute('w') ?? '', 10);
	if (Number.isFinite(width)) line.widthEmu = width;
	const fill = parseDrawingFill(ln);
	if (fill) line.fill = fill;
	const dash = first(ln, 'prstDash', NS.a)?.getAttribute('val');
	if (dash) line.dash = dash;
	const cap = ln.getAttribute('cap');
	if (cap) line.cap = cap;
	const compound = ln.getAttribute('cmpd');
	if (compound !== null) line.compound = compound;
	const alignment = ln.getAttribute('algn');
	if (alignment !== null) line.alignment = alignment;
	const custom = first(ln, 'custDash', NS.a);
	if (custom && !dash) {
		line.customDash = children(custom, 'ds', NS.a).flatMap((node) => {
			const dash = drawingDashLength(node.getAttribute('d'));
			const space = drawingDashLength(node.getAttribute('sp'));
			return dash !== undefined && space !== undefined
				? [{ dash, space, sourceXml: buildXml(node) }]
				: [];
		});
	}
	for (const kind of ['round', 'bevel', 'miter'] as const) {
		const join = first(ln, kind, NS.a);
		if (!join) continue;
		const limit = Number(join.getAttribute('lim'));
		line.join = { kind, ...(join.hasAttribute('lim') && Number.isFinite(limit) ? { limit } : {}) };
		break;
	}
	for (const end of ['headEnd', 'tailEnd'] as const) {
		const node = first(ln, end, NS.a);
		if (!node) continue;
		line[end] = {};
		for (const [attribute, key] of [
			['type', 'type'],
			['w', 'width'],
			['len', 'length'],
		] as const) {
			const value = node.getAttribute(attribute);
			if (value !== null) line[end][key] = value;
		}
	}
	return line;
}

/** ST_PositivePercentage custom-dash bounds, matching the PowerPoint object-tree reader. */
export function drawingDashLength(value: string | null): number | undefined {
	if (!/^\d+$/u.test(value ?? '')) return undefined;
	const number = Number(value);
	return Number.isSafeInteger(number) && number <= 2147483647 ? number : undefined;
}
