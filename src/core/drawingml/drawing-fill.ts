import { parseDrawingPictureFill } from './picture-fill';
// DrawingML fill and line readers over the DOM.
import { parseDrawingColorIn } from './drawing-color';
import { NS, elements, first, type XmlElement } from './dom';
import { readDrawingGradient } from './gradient-read';
import { readDrawingLineFields } from './line-read';
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

export function isDrawingFillElement(element: XmlElement): boolean {
	return (
		FILL_ELEMENTS.has(element.localName) && (!element.namespaceURI || element.namespaceURI === NS.a)
	);
}

/** The fill element held directly by a shape-properties or line element, if any. */
export function fillElementOf(properties: XmlElement | undefined): XmlElement | undefined {
	return properties ? elements(properties).find(isDrawingFillElement) : undefined;
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

/** Reads `a:ln`: the shared line fields (`line-read.ts`), the fill and the source XML. */
export function parseDrawingLine(ln: XmlElement | undefined): DrawingLine | undefined {
	if (!ln) return undefined;
	const { widthEmu, ...fields } = readDrawingLineFields(DOM_DRAWING_VIEW, ln, (node) =>
		buildXml(node),
	);
	const fill = parseDrawingFill(ln);
	return {
		sourceXml: buildXml(ln),
		...(widthEmu !== undefined ? { widthEmu } : {}),
		...(fill ? { fill } : {}),
		...fields,
	};
}

export { drawingDashLength } from './line-read';
