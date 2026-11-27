// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes a text box's changed paragraphs and tables (`InlineImage.textBoxContent.blocks`) into
// its `w:txbxContent` through the body's block writer, so a merged field result inside a text box
// is saved. Only the blocks that changed are rewritten; the shape keeps its XML.
import type { InlineImage } from './model';
import type { RelationshipAllocator } from './relationship-allocator';
import { applyBlocks } from './write';
import { W_NS, WPS_NS } from './shape-parse';
import { isElement, type XmlDocument, type XmlElement } from './xml';

/** Whether the text box content of `image` differs from the loaded `base`. */
export function textBoxBlocksChanged(image: InlineImage, base: InlineImage): boolean {
	const next = image.textBoxContent?.blocks;
	const previous = base.textBoxContent?.blocks;
	if (!next || !previous || next === previous || image.textBoxEditable) return false;
	return JSON.stringify(next) !== JSON.stringify(previous);
}

function textBoxContainer(drawing: XmlElement): XmlElement | undefined {
	const box = drawing.getElementsByTagNameNS(WPS_NS, 'txbx')[0];
	for (const child of Array.from(box?.childNodes ?? []))
		if (isElement(child) && child.namespaceURI === W_NS && child.localName === 'txbxContent')
			return child;
	return undefined;
}

/**
 * Rewrites the changed blocks of a DrawingML text box. Its VML fallback (an older copy of the same
 * text) is removed rather than left stale. Returns false when the drawing has no text box content.
 */
export function patchTextBoxBlocks(
	doc: XmlDocument,
	drawing: XmlElement,
	container: XmlElement | undefined,
	image: InlineImage,
	base: InlineImage,
	allocator: RelationshipAllocator | undefined,
): boolean {
	const content = textBoxContainer(drawing);
	const blocks = image.textBoxContent?.blocks;
	const original = base.textBoxContent?.blocks;
	if (!content || !blocks || !original) return false;
	applyBlocks(doc, content, blocks, original, allocator, Math.round(image.widthPx * 15));
	if (container?.localName === 'AlternateContent')
		for (const child of Array.from(container.childNodes))
			if (isElement(child) && child.localName === 'Fallback') container.removeChild(child);
	return true;
}
