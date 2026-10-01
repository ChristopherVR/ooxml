// DrawingML fill and line readers over the DOM. Diagram-independent; destined for `drawingml`.
import { parseDrawingColorIn } from './drawing-color.js';
import { NS, children, elements, first, type XmlElement } from './dom.js';
import type { DiagramFill, DiagramLine } from './types.js';

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
export function parseDrawingFill(properties: XmlElement | undefined): DiagramFill | undefined {
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
			const stops = children(first(fill, 'gsLst', NS.a) ?? fill, 'gs', NS.a).flatMap((stop) => {
				const color = parseDrawingColorIn(stop);
				const position = Number.parseInt(stop.getAttribute('pos') ?? '', 10);
				return color && Number.isFinite(position) ? [{ position: position / 1000, color }] : [];
			});
			const angle = Number.parseInt(first(fill, 'lin', NS.a)?.getAttribute('ang') ?? '', 10);
			const path = first(fill, 'path', NS.a)?.getAttribute('path');
			return {
				kind: 'gradient',
				stops,
				...(Number.isFinite(angle) ? { angle: angle / 60000 } : {}),
				...(path ? { path } : {}),
			};
		}
		case 'pattFill': {
			const foreground = parseDrawingColorIn(first(fill, 'fgClr', NS.a));
			const background = parseDrawingColorIn(first(fill, 'bgClr', NS.a));
			return {
				kind: 'pattern',
				preset: fill.getAttribute('prst') ?? '',
				...(foreground ? { foreground } : {}),
				...(background ? { background } : {}),
			};
		}
		case 'blipFill': {
			const blip = first(fill, 'blip', NS.a);
			const relId =
				blip?.getAttributeNS(NS.r, 'embed') || blip?.getAttribute('r:embed') || undefined;
			return { kind: 'picture', ...(relId ? { relId } : {}) };
		}
		default:
			return { kind: 'unsupported', element: fill.localName };
	}
}

/** Reads `a:ln`: width in EMU, fill, dash and cap. */
export function parseDrawingLine(ln: XmlElement | undefined): DiagramLine | undefined {
	if (!ln) return undefined;
	const line: DiagramLine = {};
	const width = Number.parseInt(ln.getAttribute('w') ?? '', 10);
	if (Number.isFinite(width)) line.widthEmu = width;
	const fill = parseDrawingFill(ln);
	if (fill) line.fill = fill;
	const dash = first(ln, 'prstDash', NS.a)?.getAttribute('val');
	if (dash) line.dash = dash;
	const cap = ln.getAttribute('cap');
	if (cap) line.cap = cap;
	return line;
}
