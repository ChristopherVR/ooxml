// DrawingML paragraph properties writer: the inverse of `read-paragraph-properties.ts`, children
// in CT_TextParagraphProperties order with the kept children slotted back in place.
import { NS, buildXml, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { drawingColorXml } from './write-color';
import { writeDrawingFont, writeDrawingRunProperties } from './write-run-properties';
import type { DrawingParagraphProperties } from './text-model';
import type { DrawingTextSpacing } from './types';
import {
	drawingBulletChild,
	drawingBulletSizeChild,
	drawingParagraphAttributeEntries,
	drawingTabStopEntries,
	drawingTextSpacingChild,
	type DrawingAttributeEntry,
} from './text-paragraph-codec';
import {
	appendInSchemaOrder,
	drawingElement,
	importDrawingXml,
	keptChildren,
	type OrderedChild,
} from './write-text-xml';

/** CT_TextParagraphProperties child order. */
export const PARAGRAPH_PROPERTIES_ORDER = [
	'lnSpc',
	'spcBef',
	'spcAft',
	'buClrTx',
	'buClr',
	'buSzTx',
	'buSzPct',
	'buSzPts',
	'buFontTx',
	'buFont',
	'buNone',
	'buAutoNum',
	'buChar',
	'buBlip',
	'tabLst',
	'defRPr',
	'extLst',
] as const;

/** An `a:<local>` element with the codec's attributes. */
function codecElement(
	doc: XmlDocument,
	local: string,
	attributes: readonly DrawingAttributeEntry[],
): XmlElement {
	const element = drawingElement(doc, local);
	for (const [name, text] of attributes) element.setAttribute(name, text);
	return element;
}

/** Writes `a:lnSpc` / `a:spcBef` / `a:spcAft`. */
export function writeDrawingTextSpacing(
	doc: XmlDocument,
	local: string,
	spacing: DrawingTextSpacing,
): XmlElement {
	const element = drawingElement(doc, local);
	const inner = drawingTextSpacingChild(spacing);
	element.appendChild(codecElement(doc, inner.local, inner.attributes));
	return element;
}

function bulletChildren(doc: XmlDocument, value: DrawingParagraphProperties): OrderedChild[] {
	const out: OrderedChild[] = [];
	const add = (name: string, node: XmlElement) => out.push({ name, node });
	if (value.bulletColor === 'text') add('buClrTx', drawingElement(doc, 'buClrTx'));
	else if (value.bulletColor) {
		const color = drawingElement(doc, 'buClr');
		color.appendChild(importDrawingXml(doc, drawingColorXml(value.bulletColor)));
		add('buClr', color);
	}
	if (value.bulletSize) {
		const size = drawingBulletSizeChild(value.bulletSize);
		add(size.local, codecElement(doc, size.local, size.attributes));
	}
	if (value.bulletFont === 'text') add('buFontTx', drawingElement(doc, 'buFontTx'));
	else if (value.bulletFont) add('buFont', writeDrawingFont(doc, 'buFont', value.bulletFont));
	const bullet = value.bullet;
	if (bullet?.kind === 'picture') add('buBlip', importDrawingXml(doc, bullet.xml));
	else if (bullet) {
		const child = drawingBulletChild(bullet);
		if (child) add(child.local, codecElement(doc, child.local, child.attributes));
	}
	return out;
}

/** Writes paragraph properties as `a:<local>` (`pPr`, `defPPr`, `lvl1pPr`...) into `doc`. */
export function writeDrawingParagraphProperties(
	doc: XmlDocument,
	value: DrawingParagraphProperties,
	local = 'pPr',
): XmlElement {
	const element = codecElement(doc, local, drawingParagraphAttributeEntries(value));
	const children: OrderedChild[] = [];
	for (const [key, name] of [
		['lineSpacing', 'lnSpc'],
		['spaceBefore', 'spcBef'],
		['spaceAfter', 'spcAft'],
	] as const) {
		const spacing = value[key];
		if (spacing) children.push({ name, node: writeDrawingTextSpacing(doc, name, spacing) });
	}
	children.push(...bulletChildren(doc, value));
	if (value.tabStops) {
		const list = drawingElement(doc, 'tabLst');
		for (const stop of value.tabStops)
			list.appendChild(codecElement(doc, 'tab', drawingTabStopEntries(stop)));
		children.push({ name: 'tabLst', node: list });
	}
	if (value.defaultRunProperties)
		children.push({
			name: 'defRPr',
			node: writeDrawingRunProperties(doc, value.defaultRunProperties, 'defRPr'),
		});
	appendInSchemaOrder(element, PARAGRAPH_PROPERTIES_ORDER, [
		...children,
		...keptChildren(doc, value.extraXml),
	]);
	return element;
}

/** Paragraph properties as an XML string (`<a:pPr .../>`). */
export function drawingParagraphPropertiesXml(
	value: DrawingParagraphProperties,
	local = 'pPr',
): string {
	const doc = parseXml(`<a:w xmlns:a="${NS.a}" xmlns:r="${NS.r}"/>`);
	return buildXml(writeDrawingParagraphProperties(doc, value, local));
}
