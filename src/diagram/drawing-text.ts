// DrawingML text body reader (`dsp:txBody`, `a:txBody`) over the DOM. Diagram-independent;
// destined for `drawingml`.
import { parseDrawingColorIn } from './drawing-color.js';
import { NS, booleanAttribute, children, first, type XmlElement } from './dom.js';
import type { DiagramTextBody, DiagramTextParagraph, DiagramTextRun } from './types.js';

function parseRun(run: XmlElement): DiagramTextRun {
	const properties = first(run, 'rPr', NS.a);
	const parsed: DiagramTextRun = { text: first(run, 't', NS.a)?.textContent ?? '' };
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

function parseParagraph(paragraph: XmlElement): DiagramTextParagraph {
	const runs: DiagramTextRun[] = [];
	for (const child of children(paragraph, 'r', NS.a)) runs.push(parseRun(child));
	// Fields (`a:fld`, slide numbers and the like) show their cached text.
	for (const field of children(paragraph, 'fld', NS.a)) runs.push(parseRun(field));
	const align = first(paragraph, 'pPr', NS.a)?.getAttribute('algn');
	return { ...(align ? { align } : {}), runs };
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
