// DrawingML paragraph properties writer: the inverse of `read-paragraph-properties.ts`, children
// in CT_TextParagraphProperties order with the kept children slotted back in place.
import { NS, buildXml, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { drawingColorXml } from './write-color';
import { writeDrawingFont, writeDrawingRunProperties } from './write-run-properties';
import type { DrawingParagraphProperties } from './text-model';
import type { DrawingTextSpacing } from './types';
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

/** Writes `a:lnSpc` / `a:spcBef` / `a:spcAft`. */
export function writeDrawingTextSpacing(
	doc: XmlDocument,
	local: string,
	spacing: DrawingTextSpacing,
): XmlElement {
	const element = drawingElement(doc, local);
	const inner = drawingElement(doc, spacing.unit === 'points' ? 'spcPts' : 'spcPct');
	setScaled(inner, 'val', spacing.value, spacing.unit === 'points' ? 100 : 100000);
	element.appendChild(inner);
	return element;
}

function setAttributes(element: XmlElement, value: DrawingParagraphProperties): void {
	setValue(element, 'marL', value.marginLeftEmu);
	setValue(element, 'marR', value.marginRightEmu);
	setValue(element, 'lvl', value.level);
	setValue(element, 'indent', value.indentEmu);
	setValue(element, 'algn', value.align);
	setValue(element, 'defTabSz', value.defaultTabSizeEmu);
	setBoolean(element, 'rtl', value.rightToLeft);
	setBoolean(element, 'eaLnBrk', value.eastAsianLineBreak);
	setValue(element, 'fontAlgn', value.fontAlign);
	setBoolean(element, 'latinLnBrk', value.latinLineBreak);
	setBoolean(element, 'hangingPunct', value.hangingPunctuation);
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
	const size = value.bulletSize;
	if (size === 'text') add('buSzTx', drawingElement(doc, 'buSzTx'));
	else if (size) {
		const local = size.unit === 'percent' ? 'buSzPct' : 'buSzPts';
		const element = drawingElement(doc, local);
		setScaled(element, 'val', size.value, size.unit === 'percent' ? 100000 : 100);
		add(local, element);
	}
	if (value.bulletFont === 'text') add('buFontTx', drawingElement(doc, 'buFontTx'));
	else if (value.bulletFont) add('buFont', writeDrawingFont(doc, 'buFont', value.bulletFont));
	const bullet = value.bullet;
	if (bullet?.kind === 'none') add('buNone', drawingElement(doc, 'buNone'));
	else if (bullet?.kind === 'char') {
		const element = drawingElement(doc, 'buChar');
		element.setAttribute('char', bullet.char);
		add('buChar', element);
	} else if (bullet?.kind === 'autoNumber') {
		const element = drawingElement(doc, 'buAutoNum');
		element.setAttribute('type', bullet.scheme);
		setValue(element, 'startAt', bullet.startAt);
		add('buAutoNum', element);
	} else if (bullet?.kind === 'picture') add('buBlip', importDrawingXml(doc, bullet.xml));
	return out;
}

/** Writes paragraph properties as `a:<local>` (`pPr`, `defPPr`, `lvl1pPr`...) into `doc`. */
export function writeDrawingParagraphProperties(
	doc: XmlDocument,
	value: DrawingParagraphProperties,
	local = 'pPr',
): XmlElement {
	const element = drawingElement(doc, local);
	setAttributes(element, value);
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
		for (const stop of value.tabStops) {
			const tab = drawingElement(doc, 'tab');
			setValue(tab, 'pos', stop.positionEmu);
			setValue(tab, 'algn', stop.align);
			list.appendChild(tab);
		}
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
