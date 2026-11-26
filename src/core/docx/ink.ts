// Word ink: a `w:drawing` whose `a:graphicData` (uri `wordprocessingInk`) holds a `w14:contentPart`
// naming an InkML part through a `customXml` relationship. Word 16 writes it bare inside the
// drawing; a pasted ink also wraps it in `mc:AlternateContent` (`Requires="a14"`) with a picture
// fallback, and the run may wrap the whole drawing in `mc:Choice Requires="wpi"`. The strokes
// are read from the part after the document is parsed, with the shared InkML reader loaded on
// first use, and are placed in the drawing's box (CSS pixels).
import type { InkStroke } from '../drawingml/inkml/types';
import { resolvePartPath } from '../opc/index';
import { forEachParagraph } from './parse-warnings';
import type { Block } from './model';
import type { Relationship } from './package-parts';
import { getR, isElement, type XmlElement } from './xml';

export const WORD_INK_GRAPHIC_URI =
	'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const INK_RELATIONSHIP =
	'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';

/** The ink of a Word drawing. */
export interface DocxInk {
	/** The `customXml` relationship naming the part; absent for ink drawn in this session. */
	relId?: string;
	/** The InkML part, e.g. `word/ink/ink1.xml`. */
	partName?: string;
	/** The strokes, in a `width` x `height` pixel space (the drawing's box when they were placed). */
	strokes: InkStroke[];
	width: number;
	height: number;
}

function contentPartOf(graphicData: XmlElement): XmlElement | undefined {
	const stack: XmlElement[] = [graphicData];
	while (stack.length) {
		const node = stack.shift()!;
		for (const child of Array.from(node.childNodes)) {
			if (!isElement(child)) continue;
			if (child.namespaceURI === W14_NS && child.localName === 'contentPart') return child;
			// Only the modern branch: a fallback holds the picture, not the ink.
			if (child.localName !== 'Fallback') stack.push(child);
		}
	}
	return undefined;
}

/** The ink of a `wordprocessingInk` graphic, with its strokes still to be read. */
export function parseInkGraphic(
	graphicData: XmlElement,
	rels: ReadonlyMap<string, Relationship>,
	widthPx: number,
	heightPx: number,
): DocxInk {
	const part = contentPartOf(graphicData);
	const relId = getR(part, 'id') || undefined;
	const rel = relId ? rels.get(relId) : undefined;
	const partName =
		rel && rel.mode !== 'External' ? resolvePartPath('word/document.xml', rel.target) : undefined;
	return {
		...(relId ? { relId } : {}),
		...(partName ? { partName } : {}),
		strokes: [],
		width: widthPx,
		height: heightPx,
	};
}

/** Every ink in `blocks`, in document order. */
export function inksIn(blocks: Block[]): DocxInk[] {
	const inks: DocxInk[] = [];
	forEachParagraph(blocks, (paragraph) => {
		for (const run of paragraph.runs) if (run.image?.ink) inks.push(run.image.ink);
	});
	return inks;
}

/** Reads the strokes of every ink in `blocks` (body, headers, footers, notes) from its part. */
export async function resolveDocumentInk(
	blocks: Block[],
	readText: (partName: string) => Promise<string | undefined>,
): Promise<void> {
	const inks = inksIn(blocks).filter((ink) => ink.partName);
	if (!inks.length) return;
	const { readInkPart } = await import('#inkml-runtime');
	for (const ink of inks) {
		const xml = await readText(ink.partName!);
		ink.strokes = readInkPart(xml, { width: ink.width, height: ink.height }).strokes;
	}
}
