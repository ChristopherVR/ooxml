// DrawingML text body reader (`dsp:txBody`, `a:txBody`) over the DOM. Diagram-independent;
// destined for `drawingml`.
import { parseDrawingColorIn } from './drawing-color';
import { NS, booleanAttribute, children, elements, first, type XmlElement } from './dom';
import type {
	DiagramTextBody,
	DiagramTextParagraph,
	DiagramTextRun,
	DiagramTextSpacing,
} from './types';

function parseSpacing(parent: XmlElement | undefined): DiagramTextSpacing | undefined {
	for (const [name, unit, divisor] of [
		['spcPts', 'points', 100],
		['spcPct', 'percent', 100000],
	] as const) {
		const raw = first(parent, name, NS.a)?.getAttribute('val');
		if (raw == null || !/^\d+$/.test(raw)) continue;
		const value = Number(raw) / divisor;
		if (Number.isFinite(value)) return { unit, value };
	}
	return undefined;
}

function parseProperties(properties: XmlElement | undefined): Omit<DiagramTextRun, 'text'> {
	const parsed: Omit<DiagramTextRun, 'text'> = {};
	const size = Number.parseInt(properties?.getAttribute('sz') ?? '', 10);
	if (Number.isFinite(size)) parsed.sizePt = size / 100;
	const bold = booleanAttribute(properties, 'b');
	if (bold !== undefined) parsed.bold = bold;
	const italic = booleanAttribute(properties, 'i');
	if (italic !== undefined) parsed.italic = italic;
	const underline = properties?.getAttribute('u');
	if (underline) parsed.underline = underline !== 'none';
	const color = parseDrawingColorIn(first(properties, 'solidFill', NS.a));
	if (color) parsed.color = color;
	const typeface = first(properties, 'latin', NS.a)?.getAttribute('typeface');
	if (typeface) parsed.typeface = typeface;
	return parsed;
}

function parseRun(run: XmlElement): DiagramTextRun {
	return {
		text: run.localName === 'br' ? '\n' : (first(run, 't', NS.a)?.textContent ?? ''),
		...parseProperties(first(run, 'rPr', NS.a)),
	};
}

function parseParagraph(paragraph: XmlElement): DiagramTextParagraph {
	const runs: DiagramTextRun[] = [];
	// Fields show cached text in their original position, and breaks retain their line boundary.
	for (const child of elements(paragraph))
		if (child.namespaceURI === NS.a && ['r', 'fld', 'br'].includes(child.localName))
			runs.push(parseRun(child));
	const properties = first(paragraph, 'pPr', NS.a);
	const align = properties?.getAttribute('algn');
	const defaults = parseProperties(first(properties, 'defRPr', NS.a));
	const lineSpacing = parseSpacing(first(properties, 'lnSpc', NS.a));
	const spaceBefore = parseSpacing(first(properties, 'spcBef', NS.a));
	const spaceAfter = parseSpacing(first(properties, 'spcAft', NS.a));
	return {
		...(align ? { align } : {}),
		...(Object.keys(defaults).length ? { defaultProperties: defaults } : {}),
		...(lineSpacing ? { lineSpacing } : {}),
		...(spaceBefore ? { spaceBefore } : {}),
		...(spaceAfter ? { spaceAfter } : {}),
		runs,
	};
}

/** Reads a text body: anchor, insets and the runs of every paragraph. */
export function parseDrawingTextBody(body: XmlElement | undefined): DiagramTextBody | undefined {
	if (!body) return undefined;
	const properties = first(body, 'bodyPr', NS.a);
	const paragraphs = children(body, 'p', NS.a).map(parseParagraph);
	const result: DiagramTextBody = {
		paragraphs,
		text: paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n'),
	};
	const anchor = properties?.getAttribute('anchor');
	if (anchor) result.anchor = anchor;
	const insets: NonNullable<DiagramTextBody['insetsEmu']> = {};
	for (const [key, attribute] of [
		['left', 'lIns'],
		['top', 'tIns'],
		['right', 'rIns'],
		['bottom', 'bIns'],
	] as const) {
		const value = Number.parseInt(properties?.getAttribute(attribute) ?? '', 10);
		if (Number.isFinite(value)) insets[key] = value;
	}
	if (Object.keys(insets).length > 0) result.insetsEmu = insets;
	return result;
}
