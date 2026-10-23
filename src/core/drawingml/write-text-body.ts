// DrawingML rich text body writer: `a:bodyPr`, `a:lstStyle` and the paragraphs, the inverse of
// `read-text-body.ts`. The container element is the caller's (`a:txBody`, `dsp:txBody`, `dgm:t`).
import { NS, buildXml, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { writeDrawingParagraphProperties } from './write-paragraph-properties';
import { writeDrawingRunProperties } from './write-run-properties';
import type {
	DrawingBodyProperties,
	DrawingListStyle,
	DrawingParagraph,
	DrawingParagraphItem,
	DrawingRichTextBody,
} from './text-model';
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

/** CT_TextBodyProperties child order. */
export const BODY_PROPERTIES_ORDER = [
	'prstTxWarp',
	'noAutofit',
	'normAutofit',
	'spAutoFit',
	'scene3d',
	'sp3d',
	'flatTx',
	'extLst',
] as const;

/** Writes `a:bodyPr`. */
export function writeDrawingBodyProperties(
	doc: XmlDocument,
	value: DrawingBodyProperties,
): XmlElement {
	const element = drawingElement(doc, 'bodyPr');
	setScaled(element, 'rot', value.rotation, 60000);
	setBoolean(element, 'spcFirstLastPara', value.spaceFirstLastParagraph);
	setValue(element, 'vertOverflow', value.verticalOverflow);
	setValue(element, 'horzOverflow', value.horizontalOverflow);
	setValue(element, 'vert', value.vertical);
	setValue(element, 'wrap', value.wrap);
	setValue(element, 'lIns', value.insetLeftEmu);
	setValue(element, 'tIns', value.insetTopEmu);
	setValue(element, 'rIns', value.insetRightEmu);
	setValue(element, 'bIns', value.insetBottomEmu);
	setValue(element, 'numCol', value.columnCount);
	setValue(element, 'spcCol', value.columnSpacingEmu);
	setBoolean(element, 'rtlCol', value.rightToLeftColumns);
	setBoolean(element, 'fromWordArt', value.fromWordArt);
	setValue(element, 'anchor', value.anchor);
	setBoolean(element, 'anchorCtr', value.anchorCenter);
	setBoolean(element, 'forceAA', value.forceAntiAlias);
	setBoolean(element, 'upright', value.upright);
	setBoolean(element, 'compatLnSpc', value.compatibleLineSpacing);
	const children: OrderedChild[] = [];
	const autofit = value.autofit;
	if (autofit) {
		const local =
			autofit.kind === 'none'
				? 'noAutofit'
				: autofit.kind === 'shape'
					? 'spAutoFit'
					: 'normAutofit';
		const node = drawingElement(doc, local);
		if (autofit.kind === 'normal') {
			setScaled(node, 'fontScale', autofit.fontScale, 100000);
			setScaled(node, 'lnSpcReduction', autofit.lineSpacingReduction, 100000);
		}
		children.push({ name: local, node });
	}
	appendInSchemaOrder(element, BODY_PROPERTIES_ORDER, [
		...children,
		...keptChildren(doc, value.extraXml),
	]);
	return element;
}

/** Writes `a:lstStyle`. */
export function writeDrawingListStyle(doc: XmlDocument, value: DrawingListStyle): XmlElement {
	const element = drawingElement(doc, 'lstStyle');
	if (value.default)
		element.appendChild(writeDrawingParagraphProperties(doc, value.default, 'defPPr'));
	value.levels.forEach((level, index) => {
		if (level)
			element.appendChild(writeDrawingParagraphProperties(doc, level, `lvl${index + 1}pPr`));
	});
	for (const kept of keptChildren(doc, value.extraXml)) element.appendChild(kept.node);
	return element;
}

function withText(doc: XmlDocument, element: XmlElement, text: string): XmlElement {
	const t = drawingElement(doc, 't');
	t.appendChild(doc.createTextNode(text));
	element.appendChild(t);
	return element;
}

function writeItem(doc: XmlDocument, item: DrawingParagraphItem): XmlElement {
	if (item.kind === 'unknown') return importDrawingXml(doc, item.xml);
	const element = drawingElement(
		doc,
		item.kind === 'run' ? 'r' : item.kind === 'break' ? 'br' : 'fld',
	);
	if (item.kind === 'field') {
		element.setAttribute('id', item.id);
		setValue(element, 'type', item.type);
	}
	if (item.properties) element.appendChild(writeDrawingRunProperties(doc, item.properties));
	if (item.kind === 'field' && item.paragraphProperties)
		element.appendChild(writeDrawingParagraphProperties(doc, item.paragraphProperties));
	return item.kind === 'break' ? element : withText(doc, element, item.text);
}

/** Writes one `a:p`. */
export function writeDrawingParagraph(doc: XmlDocument, paragraph: DrawingParagraph): XmlElement {
	const element = drawingElement(doc, 'p');
	if (paragraph.properties)
		element.appendChild(writeDrawingParagraphProperties(doc, paragraph.properties));
	for (const item of paragraph.items) element.appendChild(writeItem(doc, item));
	if (paragraph.endProperties)
		element.appendChild(writeDrawingRunProperties(doc, paragraph.endProperties, 'endParaRPr'));
	return element;
}

/**
 * Writes the children of a text body (`a:bodyPr`, `a:lstStyle`, paragraphs) into `container`.
 * A body needs at least one paragraph (CT_TextBody), so an empty list writes one empty `a:p`.
 */
export function writeDrawingRichTextBody(container: XmlElement, body: DrawingRichTextBody): void {
	const doc = container.ownerDocument;
	if (body.bodyProperties)
		container.appendChild(writeDrawingBodyProperties(doc, body.bodyProperties));
	if (body.listStyle) container.appendChild(writeDrawingListStyle(doc, body.listStyle));
	const paragraphs = body.paragraphs.length ? body.paragraphs : [{ items: [] }];
	for (const paragraph of paragraphs) container.appendChild(writeDrawingParagraph(doc, paragraph));
}

/** A text body as XML, in a container named `qualifiedName` in namespace `namespace`. */
export function drawingRichTextBodyXml(
	body: DrawingRichTextBody,
	qualifiedName = 'a:txBody',
	namespace: string = NS.a,
): string {
	const prefix = qualifiedName.includes(':') ? qualifiedName.split(':')[0] : undefined;
	const declaration = prefix && prefix !== 'a' ? ` xmlns:${prefix}="${namespace}"` : '';
	const doc = parseXml(`<${qualifiedName}${declaration} xmlns:a="${NS.a}" xmlns:r="${NS.r}"/>`);
	writeDrawingRichTextBody(doc.documentElement, body);
	return buildXml(doc);
}
