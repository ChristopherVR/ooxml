// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The content of a newly inserted text box or shape: its model blocks (`textBoxContent.blocks`)
// written into `w:txbxContent` by the body's block writer, and the text body settings the model
// holds (insets, anchor, `wrap="none"`, autofit) written onto `wps:bodyPr`.
import type { InlineImage } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import { A_NS, W_NS, WPS_NS } from './shape-parse';
import { EMU_PER_PIXEL } from './units';
import { applyBlocks } from './write';
import { isElement, makeNS, type XmlDocument, type XmlElement } from './xml';

/** EMU per twip (ECMA-376 Part 1, 20.1.2.1). */
const EMU_PER_TWIP = 635;
/** Word's default left and right text box insets, 0.1 inch. */
const DEFAULT_SIDE_INSET_EMU = 91440;
const AUTOFIT = { shape: 'a:spAutoFit', normal: 'a:normAutofit', none: 'a:noAutofit' } as const;
const AUTOFIT_NAMES = new Set(['spAutoFit', 'normAutofit', 'noAutofit']);

const firstNs = (root: XmlElement, ns: string, local: string) =>
	root.getElementsByTagNameNS(ns, local)[0] as XmlElement | undefined;

/** Writes `bodyPr` attributes from the model's text box settings, leaving the others as they are. */
function applyBodySettings(
	doc: XmlDocument,
	body: XmlElement,
	content: NonNullable<InlineImage['textBoxContent']>,
): void {
	const insets = [
		['lIns', content.insetLeftEmu],
		['tIns', content.insetTopEmu],
		['rIns', content.insetRightEmu],
		['bIns', content.insetBottomEmu],
	] as const;
	for (const [name, value] of insets)
		if (value !== undefined) body.setAttribute(name, String(Math.round(value)));
	if (content.anchor) body.setAttribute('anchor', content.anchor);
	if (content.vertical) body.setAttribute('vert', content.vertical);
	if (content.wrap === 'none') body.setAttribute('wrap', 'none');
	if (content.autofit) {
		for (const child of Array.from(body.childNodes))
			if (isElement(child) && child.namespaceURI === A_NS && AUTOFIT_NAMES.has(child.localName))
				body.removeChild(child);
		// CT_TextBodyProperties: prstTxWarp precedes the autofit choice.
		const warp = firstNs(body, A_NS, 'prstTxWarp');
		body.insertBefore(
			makeNS(doc, A_NS, AUTOFIT[content.autofit]),
			warp ? warp.nextSibling : body.firstChild,
		);
	}
}

/**
 * Whether the box's blocks are only `lines` as plain paragraphs of plain runs, the form the simple
 * text box writer (`write-text-box.ts`) already writes and keeps editable.
 */
export function blocksArePlainLines(image: InlineImage): boolean {
	const blocks = image.textBoxContent?.blocks ?? [];
	const lines = image.textBoxText ?? [];
	return (
		blocks.every((block, index) => {
			if (block.type !== 'paragraph') return false;
			const { type: _type, id: _id, runs, ...rest } = block;
			const text = runs.map((run) => run.text).join('');
			return (
				!Object.keys(rest).length &&
				runs.every((run) => Object.keys(run).length === 1 && typeof run.text === 'string') &&
				text === (lines[index] ?? '')
			);
		}) &&
		(blocks.length === lines.length || (!lines.length && blocks.length <= 1))
	);
}

/**
 * Fills the `wps:txbx` of a new drawing's shape with `image.textBoxContent.blocks`, when the model
 * has any, replacing the plain lines the drawing writers put there. A group or a box without
 * blocks (a later box of a linked chain) is left as written.
 */
export function writeTextBoxBlocks(
	doc: XmlDocument,
	drawing: XmlElement,
	image: InlineImage,
	allocator: RelationshipAllocator | undefined,
): void {
	const content = image.textBoxContent;
	if (!content?.blocks.length || image.shape?.groupChildren) return;
	const wsp = firstNs(drawing, WPS_NS, 'wsp');
	const body = wsp && firstNs(wsp, WPS_NS, 'bodyPr');
	if (!wsp || !body) return;
	let box = firstNs(wsp, WPS_NS, 'txbx');
	if (!box) {
		box = makeNS(doc, WPS_NS, 'wps:txbx');
		wsp.insertBefore(box, body);
	}
	let story = firstNs(box, W_NS, 'txbxContent');
	if (!story) story = box.appendChild(makeNS(doc, W_NS, 'w:txbxContent')) as XmlElement;
	for (const child of Array.from(story.childNodes)) story.removeChild(child);
	const sides =
		(content.insetLeftEmu ?? DEFAULT_SIDE_INSET_EMU) +
		(content.insetRightEmu ?? DEFAULT_SIDE_INSET_EMU);
	const widthTwips = Math.max(
		1,
		Math.round((image.widthPx * EMU_PER_PIXEL - sides) / EMU_PER_TWIP),
	);
	applyBlocks(doc, story, content.blocks, [], allocator, widthTwips);
	applyBodySettings(doc, body, content);
}
