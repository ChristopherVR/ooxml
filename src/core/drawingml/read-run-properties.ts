// DrawingML run properties reader (`a:rPr`, `a:defRPr`, `a:endParaRPr`, CT_TextCharacterProperties)
// over the shared DOM. Semantics follow the PowerPoint reader this replaces
// (`pptx/core/utils/text-run-properties-parser.ts`): tolerant booleans, an authored `baseline="0"`
// kept, typefaces trimmed; values stay in the file's units.
import { parseOoxmlPercent } from '../color/color-primitives';
import { parseDrawingColorIn } from './drawing-color';
import { fillElementOf, parseDrawingFill, parseDrawingLine } from './drawing-fill';
import {
	NS,
	attributeReader,
	booleanAttribute,
	elements,
	integerAttribute,
	stringAttribute,
	type XmlElement,
} from './dom';
import type { DrawingFont, DrawingHyperlink, DrawingRunProperties } from './text-model';
import { buildXml } from '../xml/index';

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
	const typeface = stringAttribute(element, 'typeface');
	if (!element || typeface === undefined) return undefined;
	const font: DrawingFont = { typeface };
	const panose = stringAttribute(element, 'panose');
	if (panose) font.panose = panose;
	const pitchFamily = integerAttribute(element, 'pitchFamily');
	if (pitchFamily !== undefined) font.pitchFamily = pitchFamily;
	const charset = integerAttribute(element, 'charset');
	if (charset !== undefined) font.charset = charset;
	return font;
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

const BOOLEAN_ATTRIBUTES = [
	['bold', 'b'],
	['italic', 'i'],
	['kumimoji', 'kumimoji'],
	['normalizeHeight', 'normalizeH'],
	['noProof', 'noProof'],
	['dirty', 'dirty'],
	['spellingError', 'err'],
	['smartTagClean', 'smtClean'],
] as const;

const STRING_ATTRIBUTES = [
	['underline', 'u'],
	['strike', 'strike'],
	['caps', 'cap'],
	['language', 'lang'],
	['altLanguage', 'altLang'],
	['bookmark', 'bmk'],
] as const;

const FONT_SLOTS = {
	latin: 'latin',
	ea: 'eastAsian',
	cs: 'complexScript',
	sym: 'symbol',
} as const;

function readAttributes(element: XmlElement, out: DrawingRunProperties): void {
	const size = integerAttribute(element, 'sz');
	if (size !== undefined) out.sizePt = size / 100;
	for (const [key, name] of BOOLEAN_ATTRIBUTES) {
		const value = booleanAttribute(element, name);
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of STRING_ATTRIBUTES) {
		const value = stringAttribute(element, name);
		if (value) out[key] = value;
	}
	const baseline = attributeReader(element)('baseline');
	const baselineFraction = baseline === undefined ? undefined : parseOoxmlPercent(baseline);
	if (baselineFraction !== undefined) out.baseline = baselineFraction;
	const spacing = integerAttribute(element, 'spc');
	if (spacing !== undefined) out.spacingPt = spacing / 100;
	const kern = integerAttribute(element, 'kern');
	if (kern !== undefined) out.kernPt = kern / 100;
	const smartTagId = integerAttribute(element, 'smtId');
	if (smartTagId !== undefined) out.smartTagId = smartTagId;
}

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
	} else if (name in FONT_SLOTS) {
		const font = parseDrawingFont(child);
		if (font) out[FONT_SLOTS[name as keyof typeof FONT_SLOTS]] = font;
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
	else if (name === 'rtl') {
		const value = booleanAttribute(child, 'val');
		out.rightToLeft = value ?? true;
	} else extra.push(outerXml(child));
}

/** Fills the shared fill writer can express; others stay verbatim in `extraXml`. */
export const WRITABLE_FILLS = new Set(['none', 'solid', 'gradient']);

/** Reads run properties; `undefined` for a missing element. */
export function parseDrawingRunProperties(
	element: XmlElement | undefined,
): DrawingRunProperties | undefined {
	if (!element) return undefined;
	const out: DrawingRunProperties = {};
	readAttributes(element, out);
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
