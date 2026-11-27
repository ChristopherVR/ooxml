// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Places rewritten paragraphs, tables and table rows in a container that holds content controls:
// each kept element keeps its own `w:sdt` wrappers, and a new one goes inside the controls it
// shares with the element before it (a repeating section around a new item), in new `w:sdt`
// wrappers for the rest of its controls.
import type { ContentControl } from './content-control-model';
import type { Block } from './model';
import { isSdt, sdtContent } from './content-control';
import {
	createContentControl,
	patchContentControl,
	unwrapContentControl,
} from './write-content-control';
import type { FlatBlock } from './parse-content-control';
import { replaceSlots } from './write-slots';
import { isElement, type XmlDocument, type XmlElement } from './xml';

/** The controls around a block, outermost first, ending with its own; none without one. */
export const blockChain = (
	block: Pick<Block, 'contentControl' | 'outerContentControls'> | undefined,
): ContentControl[] =>
	block?.contentControl ? [...(block.outerContentControls ?? []), block.contentControl] : [];

/** The ancestor of the node that is a direct child of the container, or the node itself. */
function topLevel(node: XmlElement, container: XmlElement): XmlElement {
	let current = node;
	while (current.parentNode && current.parentNode !== container)
		current = current.parentNode as XmlElement;
	return current;
}

/** The `w:sdt` elements between `node` and `container`, outermost first. */
function sdtAncestors(node: XmlElement, container: XmlElement): XmlElement[] {
	const found: XmlElement[] = [];
	for (let parent = node.parentNode; parent && parent !== container; parent = parent.parentNode)
		if (isElement(parent) && isSdt(parent)) found.unshift(parent);
	return found;
}

const hasElementChild = (node: XmlElement): boolean => {
	for (let child = node.firstChild; child; child = child.nextSibling)
		if (child.nodeType === 1) return true;
	return false;
};

/** Drops sdt wrappers left without content after their blocks were deleted. */
function pruneEmptyControls(flat: FlatBlock[]): void {
	for (const entry of flat) {
		let sdt: XmlElement | undefined = entry.sdt;
		while (sdt) {
			const content = sdtContent(sdt);
			const parent = sdt.parentNode as XmlElement | null;
			if (!parent || (content && hasElementChild(content))) break;
			parent.removeChild(sdt);
			const outer = parent.parentNode as XmlElement | null;
			sdt = isSdt(outer) ? outer : undefined;
		}
	}
}

/** Wraps `node` where it is in a new `w:sdt` for each control, the first outermost. */
function wrap(doc: XmlDocument, node: XmlElement, controls: readonly ContentControl[]): void {
	for (const control of controls) {
		const { sdt, content } = createContentControl(doc, control);
		node.parentNode!.insertBefore(sdt, node);
		content.appendChild(node);
	}
}

/**
 * Places elements already written to `nodes` in `container`, each inside the controls of
 * `chains` (parallel to `nodes`, outermost first). Returns false, changing nothing, when the kept
 * elements changed their relative order (the caller then replaces slot by slot).
 */
export function placeControlledById(
	doc: XmlDocument,
	container: XmlElement,
	flat: FlatBlock[],
	chains: readonly (readonly ContentControl[])[],
	nodes: XmlElement[],
	anchor: XmlElement | null,
): boolean {
	const oldSet = new Set(flat.map((entry) => entry.element));
	const outputSet = new Set(nodes);
	const keptInOrder = nodes.filter((node) => oldSet.has(node));
	const documentOrder = flat.map((entry) => entry.element).filter((node) => outputSet.has(node));
	if (keptInOrder.some((node, index) => node !== documentOrder[index])) return false;
	for (const entry of flat)
		if (!outputSet.has(entry.element)) entry.element.parentNode?.removeChild(entry.element);
	pruneEmptyControls(flat);
	let previous: { node: XmlElement; chain: readonly ContentControl[] } | undefined;
	nodes.forEach((node, index) => {
		const chain = chains[index] ?? [];
		if (oldSet.has(node)) {
			previous = { node, chain };
			return;
		}
		if (previous) {
			const sdts = sdtAncestors(previous.node, container);
			const before = previous.chain;
			let depth = 0;
			if (sdts.length === before.length)
				while (
					depth < Math.min(chain.length, before.length) &&
					before[depth]!.id === chain[depth]!.id
				)
					depth++;
			else if (before.length && before.at(-1)!.id === chain.at(-1)?.id) depth = sdts.length;
			const reference = depth >= sdts.length ? previous.node : sdts[depth]!;
			(reference.parentNode as XmlElement).insertBefore(node, reference.nextSibling);
			if (sdts.length === before.length || depth < sdts.length) wrap(doc, node, chain.slice(depth));
		} else {
			// No kept element comes before this one, so the next kept one is the first of them.
			const next = keptInOrder[0];
			if (next) topLevel(next, container).parentNode!.insertBefore(node, topLevel(next, container));
			else container.insertBefore(node, anchor);
			wrap(doc, node, chain);
		}
		previous = { node, chain };
	});
	syncControls(doc, flat, chains, nodes);
	return true;
}

/** `placeControlledById` for the paragraphs and tables of a story or cell. */
export const placeBlocksById = (
	doc: XmlDocument,
	container: XmlElement,
	flat: FlatBlock[],
	blocks: readonly Block[],
	nodes: XmlElement[],
	anchor: XmlElement | null,
): boolean => placeControlledById(doc, container, flat, blocks.map(blockChain), nodes, anchor);

/**
 * Brings the `w:sdt` wrappers of `flat` (the source elements, before the edit) in step with the
 * model: edited control properties reach the `w:sdt` that already wraps an element, and a control
 * the model took off every element it held is unwrapped (a temporary control after an edit, a
 * flattened one). `nodes` are the written elements, parallel to `chains`.
 */
export function syncControls(
	doc: XmlDocument,
	flat: FlatBlock[],
	chains: readonly (readonly ContentControl[])[],
	nodes: readonly XmlElement[],
): void {
	const bySdt = new Map<XmlElement, ContentControl>();
	const dropped = new Map<XmlElement, boolean>();
	const indexOf = new Map(nodes.map((node, index) => [node, index]));
	const note = (sdt: XmlElement, control: ContentControl | undefined) => {
		if (control && !bySdt.has(sdt)) bySdt.set(sdt, control);
		dropped.set(sdt, (dropped.get(sdt) ?? true) && !control);
	};
	flat.forEach((entry) => {
		const index = indexOf.get(entry.element);
		const chain = index === undefined ? undefined : chains[index];
		// A deleted element has no say: the ones that remain decide whether their control stays.
		if (!entry.sdt || !chain) return;
		// Each `w:sdt` around it follows the model's control with its id: patched when the model
		// keeps it, unwrapped when the model took it off every element it held. A model that gives
		// the innermost one another id (and no outer one) patches it, as one control.
		const ids = new Set(chain.map((control) => control.id));
		const levels = [...(entry.outer ?? []), { control: entry.control!, sdt: entry.sdt }];
		levels.forEach((level) => {
			const own = chain.find((control) => control.id === level.control.id);
			const renamed = levels.length === 1 && chain.length === 1 && !ids.has(level.control.id);
			note(level.sdt, own ?? (renamed ? chain[0] : undefined));
		});
	});
	for (const [sdt, control] of bySdt) patchContentControl(doc, sdt, control);
	for (const [sdt, gone] of dropped) if (gone) unwrapContentControl(sdt);
}

/** `syncControls` for the paragraphs and tables of a story or cell. */
export const syncBlockControls = (
	doc: XmlDocument,
	flat: FlatBlock[],
	blocks: readonly Block[],
	nodes: readonly XmlElement[],
): void => syncControls(doc, flat, blocks.map(blockChain), nodes);

/**
 * Puts the written paragraphs of a cell in place: through `placeBlocksById` when a block control
 * is involved (so that a control's new paragraphs stay inside it), else slot by slot.
 */
export function placeCellParagraphs(
	doc: XmlDocument,
	tc: XmlElement,
	flat: FlatBlock[],
	paragraphs: readonly Block[],
	nodes: XmlElement[],
): void {
	const controlled = flat.some((entry) => entry.sdt) || paragraphs.some((p) => p.contentControl);
	if (controlled && placeBlocksById(doc, tc, flat, paragraphs, nodes, null)) return;
	replaceSlots(
		tc,
		flat.map((entry) => entry.element),
		nodes,
	);
	syncBlockControls(doc, flat, paragraphs, nodes);
}

/** Wrap consecutive authored blocks, retaining the source wrappers that already hold blocks. */
export function wrapNewBlockControls(
	doc: XmlDocument,
	blocks: readonly Block[],
	nodes: readonly XmlElement[],
): void {
	let active: (ReturnType<typeof createContentControl> & { id: string }) | undefined;
	for (const [index, node] of nodes.entries()) {
		const control = blocks[index]?.contentControl;
		const parent = node.parentNode;
		const enclosing = parent?.parentNode;
		if (!control || !parent || (enclosing && isElement(enclosing) && isSdt(enclosing))) {
			active = undefined;
			continue;
		}
		if (active?.id !== control.id || active.sdt.nextSibling !== node) {
			active = { ...createContentControl(doc, control), id: control.id };
			parent.insertBefore(active.sdt, node);
		}
		active.content.appendChild(node);
	}
}
