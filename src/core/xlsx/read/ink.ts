import { INK_RELATIONSHIP_TYPE } from '../../drawingml/inkml/constants';
import { NS, elements, first, relAttr, type XmlElement } from '../../xml/index';
import type { DrawingAnchor, InkObject, Worksheet } from '../model';
import type { SourceIndex } from './package';
import { att, selfContainedXml } from './xml-util';

/** Excel 2010's drawing namespace, which qualifies the children of an ink `xdr:contentPart`. */
export const XDR14_NS = 'http://schemas.microsoft.com/office/excel/2010/spreadsheetDrawing';
const EMU_PER_PIXEL = 9525;

/**
 * The `contentPart` of an anchor's content: the content itself (Excel's save without a fallback),
 * or the one inside its `mc:AlternateContent` choice (a pasted ink with a fallback picture).
 */
function contentPartOf(content: XmlElement): XmlElement | undefined {
	if (content.localName === 'contentPart') return content;
	if (content.localName !== 'AlternateContent') return undefined;
	for (const choice of elements(content))
		if (choice.localName === 'Choice') {
			const part = elements(choice).find((node) => node.localName === 'contentPart');
			if (part) return part;
		}
	return undefined;
}

const child = (parent: XmlElement | undefined, local: string) =>
	parent ? elements(parent).find((node) => node.localName === local) : undefined;

/**
 * Reads an ink drawing (Excel writes `xdr:contentPart` with `xdr14:` children, sometimes inside
 * `mc:AlternateContent` with a picture fallback), or undefined when `content` is not ink.
 */
export function readInkFrame(
	source: SourceIndex,
	partName: string,
	anchorEl: XmlElement,
	content: XmlElement,
	anchor: DrawingAnchor,
): InkObject | null | undefined {
	const part = contentPartOf(content);
	if (!part) return undefined;
	const rel = source.rels(partName).get(relAttr(part, 'id') ?? '');
	const target =
		rel && rel.mode !== 'External' && rel.type === INK_RELATIONSHIP_TYPE
			? source.target(partName, rel)
			: undefined;
	// Without its part the drawing cannot be read: the caller keeps it, verbatim, as an object
	// not shown.
	if (!target) return null;
	const nv = child(part, 'nvContentPartPr');
	const cNvPr = child(nv, 'cNvPr');
	const ext = first(child(part, 'xfrm'), 'ext', NS.a);
	const width = Number(att(ext, 'cx') ?? 0) / EMU_PER_PIXEL || 96;
	const height = Number(att(ext, 'cy') ?? 0) / EMU_PER_PIXEL || 96;
	const ink: InkObject = {
		kind: 'ink',
		anchor,
		// Read from the part after the drawings, by `resolveWorkbookInk`.
		strokes: [],
		width,
		height,
		sourceXml: selfContainedXml(anchorEl),
	};
	ink.partName = target;
	const name = att(cNvPr, 'name');
	if (name) ink.name = name;
	const description = att(cNvPr, 'descr');
	if (description) ink.description = description;
	return ink;
}

/** Reads the strokes of every ink drawing from its part, with the InkML reader loaded on first use. */
export async function resolveWorkbookInk(
	source: SourceIndex,
	sheets: readonly Worksheet[],
): Promise<void> {
	const inks = sheets.flatMap((sheet) =>
		sheet.drawings.filter(
			(drawing): drawing is InkObject => drawing.kind === 'ink' && Boolean(drawing.partName),
		),
	);
	if (!inks.length) return;
	const { readInkPart } = await import('#inkml-runtime');
	for (const ink of inks) {
		const part = readInkPart(source.text(ink.partName!), { width: ink.width, height: ink.height });
		ink.strokes = part.strokes;
		if (part.truncated) ink.truncated = true;
	}
}
