// DrawingML rich text body reader: `a:bodyPr` (CT_TextBodyProperties), `a:lstStyle` and the
// paragraphs with their ordered runs, line breaks and fields, over the shared DOM. Works for any
// element holding a text body (`a:txBody`, `p:txBody`, `dsp:txBody`, `dgm:t`, `c:rich`, `c:txPr`).
import { NS, attributeReader, booleanAttribute, children, elements, first } from './dom';
import { integerAttribute, stringAttribute, type XmlElement } from './dom';
import { parseDrawingParagraphProperties } from './read-paragraph-properties';
import { isDrawingElement, outerXml, parseDrawingRunProperties } from './read-run-properties';
import type {
	DrawingBodyProperties,
	DrawingListStyle,
	DrawingParagraph,
	DrawingParagraphItem,
	DrawingRichTextBody,
} from './text-model';
import { parseOoxmlPercent } from '../color/color-primitives';

const BODY_STRINGS = [
	['verticalOverflow', 'vertOverflow'],
	['horizontalOverflow', 'horzOverflow'],
	['vertical', 'vert'],
	['wrap', 'wrap'],
	['anchor', 'anchor'],
] as const;

const BODY_INTEGERS = [
	['insetLeftEmu', 'lIns'],
	['insetTopEmu', 'tIns'],
	['insetRightEmu', 'rIns'],
	['insetBottomEmu', 'bIns'],
	['columnCount', 'numCol'],
	['columnSpacingEmu', 'spcCol'],
] as const;

const BODY_BOOLEANS = [
	['spaceFirstLastParagraph', 'spcFirstLastPara'],
	['rightToLeftColumns', 'rtlCol'],
	['fromWordArt', 'fromWordArt'],
	['anchorCenter', 'anchorCtr'],
	['forceAntiAlias', 'forceAA'],
	['upright', 'upright'],
	['compatibleLineSpacing', 'compatLnSpc'],
] as const;

/** Reads `a:bodyPr`; `undefined` for a missing element. */
export function parseDrawingBodyProperties(
	element: XmlElement | undefined,
): DrawingBodyProperties | undefined {
	if (!element) return undefined;
	const out: DrawingBodyProperties = {};
	const rotation = integerAttribute(element, 'rot');
	if (rotation !== undefined) out.rotation = rotation / 60000;
	for (const [key, name] of BODY_STRINGS) {
		const value = stringAttribute(element, name);
		if (value) out[key] = value;
	}
	for (const [key, name] of BODY_INTEGERS) {
		const value = integerAttribute(element, name);
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of BODY_BOOLEANS) {
		const value = booleanAttribute(element, name);
		if (value !== undefined) out[key] = value;
	}
	const extra: string[] = [];
	for (const child of elements(element)) {
		const name = isDrawingElement(child) ? child.localName : '';
		if (name === 'noAutofit') out.autofit = { kind: 'none' };
		else if (name === 'spAutoFit') out.autofit = { kind: 'shape' };
		else if (name === 'normAutofit') {
			const read = attributeReader(child);
			const fontScale = parseOoxmlPercent(read('fontScale') ?? '');
			const reduction = parseOoxmlPercent(read('lnSpcReduction') ?? '');
			out.autofit = {
				kind: 'normal',
				...(read('fontScale') !== undefined && fontScale !== undefined ? { fontScale } : {}),
				...(read('lnSpcReduction') !== undefined && reduction !== undefined
					? { lineSpacingReduction: reduction }
					: {}),
			};
		} else extra.push(outerXml(child));
	}
	if (extra.length) out.extraXml = extra;
	return out;
}

/** Reads `a:lstStyle`; `undefined` for a missing element. */
export function parseDrawingListStyle(
	element: XmlElement | undefined,
): DrawingListStyle | undefined {
	if (!element) return undefined;
	const style: DrawingListStyle = { levels: [] };
	const extra: string[] = [];
	for (const child of elements(element)) {
		const level = isDrawingElement(child) ? /^lvl([1-9])pPr$/u.exec(child.localName) : null;
		if (level) style.levels[Number(level[1]) - 1] = parseDrawingParagraphProperties(child);
		else if (isDrawingElement(child) && child.localName === 'defPPr')
			style.default = parseDrawingParagraphProperties(child) ?? {};
		else extra.push(outerXml(child));
	}
	// Fill holes so `levels` is a dense array of `undefined` rather than a sparse one.
	style.levels = Array.from(style.levels);
	if (extra.length) style.extraXml = extra;
	return style;
}

function textOf(element: XmlElement): string {
	return first(element, 't', NS.a)?.textContent ?? '';
}

function readItem(child: XmlElement): DrawingParagraphItem | undefined {
	const name = isDrawingElement(child) ? child.localName : '';
	if (name === 'pPr' || name === 'endParaRPr') return undefined;
	if (name === 'r') {
		const properties = parseDrawingRunProperties(first(child, 'rPr', NS.a));
		return { kind: 'run', text: textOf(child), ...(properties ? { properties } : {}) };
	}
	if (name === 'br') {
		const properties = parseDrawingRunProperties(first(child, 'rPr', NS.a));
		return { kind: 'break', ...(properties ? { properties } : {}) };
	}
	if (name === 'fld') {
		const properties = parseDrawingRunProperties(first(child, 'rPr', NS.a));
		const paragraphProperties = parseDrawingParagraphProperties(first(child, 'pPr', NS.a));
		const type = child.getAttribute('type');
		return {
			kind: 'field',
			id: child.getAttribute('id') ?? '',
			...(type ? { type } : {}),
			text: textOf(child),
			...(properties ? { properties } : {}),
			...(paragraphProperties ? { paragraphProperties } : {}),
		};
	}
	return { kind: 'unknown', xml: outerXml(child) };
}

/** Reads one `a:p`. */
export function parseDrawingParagraph(element: XmlElement): DrawingParagraph {
	const items: DrawingParagraphItem[] = [];
	for (const child of elements(element)) {
		const item = readItem(child);
		if (item) items.push(item);
	}
	const properties = parseDrawingParagraphProperties(first(element, 'pPr', NS.a));
	const endProperties = parseDrawingRunProperties(first(element, 'endParaRPr', NS.a));
	return {
		...(properties ? { properties } : {}),
		items,
		...(endProperties ? { endProperties } : {}),
	};
}

/** Reads a full text body; `undefined` for a missing element. */
export function parseDrawingRichTextBody(
	element: XmlElement | undefined,
): DrawingRichTextBody | undefined {
	if (!element) return undefined;
	const bodyProperties = parseDrawingBodyProperties(first(element, 'bodyPr', NS.a));
	const listStyle = parseDrawingListStyle(first(element, 'lstStyle', NS.a));
	return {
		...(bodyProperties ? { bodyProperties } : {}),
		...(listStyle ? { listStyle } : {}),
		paragraphs: children(element, 'p', NS.a).map(parseDrawingParagraph),
	};
}

/** The plain text of a paragraph: runs and fields, `\n` for line breaks. */
export function drawingParagraphText(paragraph: DrawingParagraph): string {
	return paragraph.items
		.map((item) =>
			item.kind === 'run' || item.kind === 'field' ? item.text : item.kind === 'break' ? '\n' : '',
		)
		.join('');
}
