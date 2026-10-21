// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Places rewritten paragraphs and tables in a story that holds block-level content controls: each
// block keeps its own `w:sdt` wrapper, and new blocks join the control of the block before them.
import type { Block } from './model';
import { isSdt, sdtContent } from './content-control';
import { patchContentControl, unwrapContentControl } from './write-content-control';
import type { FlatBlock } from './parse-content-control';
import type { XmlDocument, XmlElement } from './xml';

/** The ancestor of the node that is a direct child of the container, or the node itself. */
function topLevel(node: XmlElement, container: XmlElement): XmlElement {
	let current = node;
	while (current.parentNode && current.parentNode !== container)
		current = current.parentNode as XmlElement;
	return current;
}

/** Drops sdt wrappers left without content after their blocks were deleted. */
function pruneEmptyControls(flat: FlatBlock[]): void {
	for (const entry of flat) {
		let sdt: XmlElement | undefined = entry.sdt;
		while (sdt) {
			const content = sdtContent(sdt);
			const parent = sdt.parentNode as XmlElement | null;
			if (!parent || (content && Array.from(content.childNodes).some((n) => n.nodeType === 1)))
				break;
			parent.removeChild(sdt);
			const outer = parent.parentNode as XmlElement | null;
			sdt = isSdt(outer) ? outer : undefined;
		}
	}
}

/**
 * Places `blocks` (already written to `nodes`) in `container`. Returns false, changing nothing,
 * when the kept blocks changed their relative order (the caller then replaces slot by slot).
 */
export function placeBlocksById(
	doc: XmlDocument,
	container: XmlElement,
	flat: FlatBlock[],
	blocks: Block[],
	nodes: XmlElement[],
	anchor: XmlElement | null,
): boolean {
	const oldSet = new Set(flat.map((entry) => entry.element));
	const keptInOrder = nodes.filter((node) => oldSet.has(node));
	const documentOrder = flat
		.map((entry) => entry.element)
		.filter((node) => keptInOrder.includes(node));
	if (keptInOrder.some((node, index) => node !== documentOrder[index])) return false;
	const outputSet = new Set(nodes);
	for (const entry of flat)
		if (!outputSet.has(entry.element)) entry.element.parentNode?.removeChild(entry.element);
	pruneEmptyControls(flat);
	let previous: { node: XmlElement; block: Block } | undefined;
	nodes.forEach((node, index) => {
		const block = blocks[index]!;
		if (oldSet.has(node)) {
			previous = { node, block };
			return;
		}
		if (previous) {
			const sameControl =
				previous.block.contentControl &&
				previous.block.contentControl.id === block.contentControl?.id;
			const reference = sameControl ? previous.node : topLevel(previous.node, container);
			(reference.parentNode as XmlElement).insertBefore(node, reference.nextSibling);
		} else {
			const next = nodes.slice(index + 1).find((candidate) => oldSet.has(candidate));
			if (next) topLevel(next, container).parentNode!.insertBefore(node, topLevel(next, container));
			else container.insertBefore(node, anchor);
		}
		previous = { node, block };
	});
	// Edited control properties reach the `w:sdt` that already wraps the block, and a control the
	// model took off every block it held is unwrapped (a temporary control after an edit).
	const bySdt = new Map<XmlElement, Block>();
	const dropped = new Map<XmlElement, boolean>();
	flat.forEach((entry) => {
		const block = blocks[nodes.indexOf(entry.element)];
		// A deleted block has no say: the blocks that remain decide whether their control stays.
		if (!entry.sdt || !block) return;
		if (block.contentControl && !bySdt.has(entry.sdt)) bySdt.set(entry.sdt, block);
		dropped.set(entry.sdt, (dropped.get(entry.sdt) ?? true) && !block.contentControl);
	});
	for (const [sdt, block] of bySdt) patchContentControl(doc, sdt, block.contentControl!);
	for (const [sdt, gone] of dropped) if (gone) unwrapContentControl(sdt);
	return true;
}
