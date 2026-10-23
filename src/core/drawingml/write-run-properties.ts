// DrawingML run properties writer: the inverse of `read-run-properties.ts`, child elements in
// CT_TextCharacterProperties order with the kept children slotted back in place.
import { NS, buildXml, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { drawingColorXml } from './write-color';
import { drawingFillXml } from './write-fill';
import { patchDrawingLine } from './write-line';
import type { DrawingFont, DrawingHyperlink, DrawingRunProperties } from './text-model';
import {
	appendInSchemaOrder,
	drawingElement,
	importDrawingXml,
	keptChildren,
	setBoolean,
	setScaled,
	setValue,
	type OrderedChild,
} from './write-text-xml';

/** CT_TextCharacterProperties child order (fills and effects by their element names). */
export const RUN_PROPERTIES_ORDER = [
	'ln',
	'noFill',
	'solidFill',
	'gradFill',
	'blipFill',
	'pattFill',
	'grpFill',
	'effectLst',
	'effectDag',
	'highlight',
	'uLnTx',
	'uLn',
	'uFillTx',
	'uFill',
	'latin',
	'ea',
	'cs',
	'sym',
	'hlinkClick',
	'hlinkMouseOver',
	'rtl',
	'extLst',
] as const;

const FILL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** Writes a `CT_TextFont` slot. */
export function writeDrawingFont(doc: XmlDocument, local: string, font: DrawingFont): XmlElement {
	const element = drawingElement(doc, local);
	element.setAttribute('typeface', font.typeface);
	setValue(element, 'panose', font.panose);
	setValue(element, 'pitchFamily', font.pitchFamily);
	setValue(element, 'charset', font.charset);
	return element;
}

function writeHyperlink(doc: XmlDocument, local: string, link: DrawingHyperlink): XmlElement {
	const element = drawingElement(doc, local);
	if (link.relId !== undefined) element.setAttributeNS(NS.r, 'r:id', link.relId);
	setValue(element, 'invalidUrl', link.invalidUrl);
	setValue(element, 'action', link.action);
	setValue(element, 'tgtFrame', link.targetFrame);
	setValue(element, 'tooltip', link.tooltip);
	setBoolean(element, 'history', link.history);
	setBoolean(element, 'highlightClick', link.highlightClick);
	setBoolean(element, 'endSnd', link.endSound);
	for (const xml of link.extraXml ?? []) element.appendChild(importDrawingXml(doc, xml));
	return element;
}

function setAttributes(element: XmlElement, value: DrawingRunProperties): void {
	setBoolean(element, 'kumimoji', value.kumimoji);
	setValue(element, 'lang', value.language);
	setValue(element, 'altLang', value.altLanguage);
	setScaled(element, 'sz', value.sizePt, 100);
	setBoolean(element, 'b', value.bold);
	setBoolean(element, 'i', value.italic);
	setValue(element, 'u', value.underline);
	setValue(element, 'strike', value.strike);
	setScaled(element, 'kern', value.kernPt, 100);
	setValue(element, 'cap', value.caps);
	setScaled(element, 'spc', value.spacingPt, 100);
	setBoolean(element, 'normalizeH', value.normalizeHeight);
	setScaled(element, 'baseline', value.baseline, 100000);
	setBoolean(element, 'noProof', value.noProof);
	setBoolean(element, 'dirty', value.dirty);
	setBoolean(element, 'err', value.spellingError);
	setBoolean(element, 'smtClean', value.smartTagClean);
	setValue(element, 'smtId', value.smartTagId);
	setValue(element, 'bmk', value.bookmark);
}

function fillChild(doc: XmlDocument, fill: DrawingRunProperties['fill']): OrderedChild | undefined {
	const xml = fill ? drawingFillXml(fill) : undefined;
	if (!xml) return undefined;
	const node = importDrawingXml(doc, xml);
	return { name: node.localName, node };
}

function modelledChildren(doc: XmlDocument, value: DrawingRunProperties): OrderedChild[] {
	const out: OrderedChild[] = [];
	const add = (name: string, node: XmlElement) => out.push({ name, node });
	if (value.line || value.lineSourceXml) {
		const ln = value.lineSourceXml
			? importDrawingXml(doc, value.lineSourceXml)
			: drawingElement(doc, 'ln');
		if (value.line) patchDrawingLine(ln, value.line);
		add('ln', ln);
	}
	const fill = fillChild(doc, value.fill);
	if (fill) out.push(fill);
	if (value.highlight) {
		const highlight = drawingElement(doc, 'highlight');
		highlight.appendChild(importDrawingXml(doc, drawingColorXml(value.highlight)));
		add('highlight', highlight);
	}
	if (value.underlineLineFollowsText) add('uLnTx', drawingElement(doc, 'uLnTx'));
	else if (value.underlineLineXml) add('uLn', importDrawingXml(doc, value.underlineLineXml));
	if (value.underlineFillFollowsText) add('uFillTx', drawingElement(doc, 'uFillTx'));
	else {
		const underline = fillChild(doc, value.underlineFill);
		if (underline) {
			const uFill = drawingElement(doc, 'uFill');
			uFill.appendChild(underline.node);
			add('uFill', uFill);
		}
	}
	for (const [local, key] of [
		['latin', 'latin'],
		['ea', 'eastAsian'],
		['cs', 'complexScript'],
		['sym', 'symbol'],
	] as const) {
		const font = value[key];
		if (font) add(local, writeDrawingFont(doc, local, font));
	}
	if (value.hyperlinkClick)
		add('hlinkClick', writeHyperlink(doc, 'hlinkClick', value.hyperlinkClick));
	if (value.hyperlinkMouseOver)
		add('hlinkMouseOver', writeHyperlink(doc, 'hlinkMouseOver', value.hyperlinkMouseOver));
	if (value.rightToLeft !== undefined) {
		const rtl = drawingElement(doc, 'rtl');
		if (!value.rightToLeft) rtl.setAttribute('val', '0');
		add('rtl', rtl);
	}
	return out;
}

/** Writes run properties as `a:<local>` (`rPr`, `defRPr`, `endParaRPr`) into `doc`. */
export function writeDrawingRunProperties(
	doc: XmlDocument,
	value: DrawingRunProperties,
	local = 'rPr',
): XmlElement {
	const element = drawingElement(doc, local);
	setAttributes(element, value);
	const modelled = modelledChildren(doc, value);
	const hasFill = modelled.some((child) => FILL_NAMES.has(child.name));
	// A modelled fill replaces a kept one (a pattern fill edited to a solid one).
	const kept = keptChildren(doc, value.extraXml).filter(
		(child) => !(hasFill && FILL_NAMES.has(child.name)),
	);
	appendInSchemaOrder(element, RUN_PROPERTIES_ORDER, [...modelled, ...kept]);
	return element;
}

/** Run properties as an XML string (`<a:rPr .../>`). */
export function drawingRunPropertiesXml(value: DrawingRunProperties, local = 'rPr'): string {
	const doc = parseXml(`<a:w xmlns:a="${NS.a}" xmlns:r="${NS.r}"/>`);
	return buildXml(writeDrawingRunProperties(doc, value, local));
}
