// DrawingML run properties reader (`a:rPr`, `a:defRPr`, `a:endParaRPr`, CT_TextCharacterProperties)
// over the shared DOM. The attribute, font slot and `a:rtl` rules are the shared run codec's
// (`text-run-codec.ts`, which PowerPoint's object-tree reader also runs): an authored
// `baseline="0"` kept, typefaces trimmed; values stay in the file's units. This reader adds the
// fills, outline, highlight and hyperlinks and keeps unknown children verbatim.
import { parseDrawingColorIn } from './drawing-color';
import { fillElementOf, parseDrawingFill, parseDrawingLine } from './drawing-fill';
import {
	NS,
	attributeReader,
	booleanAttribute,
	elements,
	stringAttribute,
	type XmlElement,
} from './dom';
import type { DrawingFont, DrawingHyperlink, DrawingRunProperties } from './text-model';
import { buildXml } from '../xml/index';
import { domTextElementView } from './text-element-view';
import { readDrawingFont } from './text-paragraph-codec';
import {
	DRAWING_RUN_FONT_SLOTS,
	readDrawingRunAttributes,
	readDrawingRunRightToLeft,
} from './text-run-codec';

/** The outer XML of an element, for children kept verbatim. */
export function outerXml(element: XmlElement): string {
	return buildXml(element);
}

/** Whether `element` is a DrawingML (`a:`) element; elements written without a namespace count. */
export function isDrawingElement(element: XmlElement): boolean {
	return !element.namespaceURI || element.namespaceURI === NS.a;
}

/** Reads a `CT_TextFont` slot; `undefined` without a typeface. */
export function parseDrawingFont(element: XmlElement | undefined): DrawingFont | undefined {
	return readDrawingFont(element ? domTextElementView(element) : undefined);
}

/** Reads `a:hlinkClick` / `a:hlinkMouseOver`. */
export function parseDrawingHyperlink(element: XmlElement): DrawingHyperlink {
	const link: DrawingHyperlink = {};
	const relId = attributeReader(element)('id')?.trim();
	if (relId) link.relId = relId;
	for (const [key, name] of [
		['tooltip', 'tooltip'],
		['action', 'action'],
		['invalidUrl', 'invalidUrl'],
		['targetFrame', 'tgtFrame'],
	] as const) {
		const value = stringAttribute(element, name);
		if (value) link[key] = value;
	}
	for (const [key, name] of [
		['history', 'history'],
		['highlightClick', 'highlightClick'],
		['endSound', 'endSnd'],
	] as const) {
		const value = booleanAttribute(element, name);
		if (value !== undefined) link[key] = value;
	}
	const extra = elements(element).map(outerXml);
	if (extra.length) link.extraXml = extra;
	return link;
}

const FONT_SLOTS: Record<string, (typeof DRAWING_RUN_FONT_SLOTS)[number][1]> =
	Object.fromEntries(DRAWING_RUN_FONT_SLOTS);

function readChild(child: XmlElement, out: DrawingRunProperties, extra: string[]): void {
	const name = child.localName;
	if (!isDrawingElement(child)) {
		extra.push(outerXml(child));
		return;
	}
	if (name === 'ln') {
		const line = parseDrawingLine(child);
		if (line) out.line = line;
		out.lineSourceXml = outerXml(child);
	} else if (Object.hasOwn(FONT_SLOTS, name)) {
		const font = parseDrawingFont(child);
		if (font) out[FONT_SLOTS[name]!] = font;
		else extra.push(outerXml(child));
	} else if (name === 'highlight') {
		const color = parseDrawingColorIn(child);
		if (color) out.highlight = color;
	} else if (name === 'uLnTx') out.underlineLineFollowsText = true;
	else if (name === 'uLn') out.underlineLineXml = outerXml(child);
	else if (name === 'uFillTx') out.underlineFillFollowsText = true;
	else if (name === 'uFill') {
		const fill = parseDrawingFill(child);
		if (fill) out.underlineFill = fill;
		else extra.push(outerXml(child));
	} else if (name === 'hlinkClick') out.hyperlinkClick = parseDrawingHyperlink(child);
	else if (name === 'hlinkMouseOver') out.hyperlinkMouseOver = parseDrawingHyperlink(child);
	else if (name === 'rtl') out.rightToLeft = readDrawingRunRightToLeft(domTextElementView(child));
	else extra.push(outerXml(child));
}

/** Fills the shared fill writer can express; others stay verbatim in `extraXml`. */
export const WRITABLE_FILLS = new Set(['none', 'solid', 'gradient']);

/** Reads run properties; `undefined` for a missing element. */
export function parseDrawingRunProperties(
	element: XmlElement | undefined,
): DrawingRunProperties | undefined {
	if (!element) return undefined;
	const out: DrawingRunProperties = {};
	readDrawingRunAttributes(domTextElementView(element), out);
	const extra: string[] = [];
	const fillElement = fillElementOf(element);
	for (const child of elements(element)) {
		if (child === fillElement) {
			const fill = parseDrawingFill(element);
			if (fill && WRITABLE_FILLS.has(fill.kind)) out.fill = fill;
			else extra.push(outerXml(child));
			continue;
		}
		readChild(child, out, extra);
	}
	if (extra.length) out.extraXml = extra;
	return out;
}
