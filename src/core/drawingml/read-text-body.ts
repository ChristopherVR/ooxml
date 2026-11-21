import {
	parseDrawingTextScene3D,
	parseDrawingTextShape3D,
	parseDrawingFlatText3D,
} from './read-text-3d';
// DrawingML rich text body reader: `a:bodyPr` (CT_TextBodyProperties), `a:lstStyle` and the
// paragraphs with their ordered runs, line breaks and fields, over the shared DOM. Works for any
// element holding a text body (`a:txBody`, `p:txBody`, `dsp:txBody`, `dgm:t`, `c:rich`, `c:txPr`).
// The `a:bodyPr` attribute and autofit rules are the shared body codec's (`text-body-codec.ts`).
import { NS, children, elements, first, type XmlElement } from './dom';
import { parseDrawingParagraphProperties } from './read-paragraph-properties';
import { isDrawingElement, outerXml, parseDrawingRunProperties } from './read-run-properties';
import {
	readDrawingAutofit,
	readDrawingBodyAttributes,
	readDrawingTextWarp,
} from './text-body-codec';
import { domTextElementView } from './text-element-view';
import type {
	DrawingBodyProperties,
	DrawingListStyle,
	DrawingParagraph,
	DrawingParagraphItem,
	DrawingRichTextBody,
} from './text-model';

/**
 * Whether the typed warp holds everything `a:prstTxWarp` says: a `prst`, no attributes besides
 * it, at most an `a:avLst` of `a:gd` guides that each have exactly `name` and `fmla`. Anything
 * else is kept verbatim.
 */
function isLosslessWarp(element: XmlElement): boolean {
	const onlyAttributes = (node: XmlElement, names: string[]) =>
		Array.from(node.attributes).every((attribute) => names.includes(attribute.name)) &&
		names.every((name) => node.hasAttribute(name));
	if (!onlyAttributes(element, ['prst']) || !element.getAttribute('prst')?.trim()) return false;
	const kids = elements(element);
	if (kids.length > 1) return false;
	const list = kids[0];
	if (!list) return true;
	if (list.localName !== 'avLst' || list.namespaceURI !== NS.a || list.attributes.length)
		return false;
	return elements(list).every(
		(guide) =>
			guide.localName === 'gd' &&
			guide.namespaceURI === NS.a &&
			onlyAttributes(guide, ['name', 'fmla']) &&
			guide.getAttribute('name') === guide.getAttribute('name')?.trim(),
	);
}

/** Reads a preset warp without interpreting the rest of the text box. */
export function parseDrawingTextWarp(element: XmlElement | undefined) {
	return element && isLosslessWarp(element)
		? readDrawingTextWarp(domTextElementView(element))
		: undefined;
}

/** Reads `a:bodyPr`; `undefined` for a missing element. */
export function parseDrawingBodyProperties(
	element: XmlElement | undefined,
): DrawingBodyProperties | undefined {
	if (!element) return undefined;
	const out: DrawingBodyProperties = {};
	readDrawingBodyAttributes(domTextElementView(element), out);
	const extra: string[] = [];
	for (const child of elements(element)) {
		const name = isDrawingElement(child) ? child.localName : '';
		const autofit = readDrawingAutofit(name, domTextElementView(child));
		if (autofit) out.autofit = autofit;
		else if (name === 'prstTxWarp' && !out.textWarp) {
			const warp = parseDrawingTextWarp(child);
			if (warp) out.textWarp = warp;
			else extra.push(outerXml(child));
		} else if (name === 'scene3d' && children(element, 'scene3d', NS.a).length === 1)
			out.scene3d = parseDrawingTextScene3D(child);
		else if (
			name === 'sp3d' &&
			children(element, 'sp3d', NS.a).length === 1 &&
			!first(element, 'flatTx', NS.a)
		)
			out.shape3d = parseDrawingTextShape3D(child);
		else if (
			name === 'flatTx' &&
			children(element, 'flatTx', NS.a).length === 1 &&
			!first(element, 'sp3d', NS.a)
		)
			out.flatText3d = parseDrawingFlatText3D(child);
		else extra.push(outerXml(child));
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
