// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Places rewritten paragraphs, tables and table rows in a container that holds content controls:
// each kept element keeps its own `w:sdt` wrappers, and a new one goes inside the controls it
// shares with the element before it (a repeating section around a new item), in new `w:sdt`
// wrappers for the rest of its controls.
import { saveEnvironmentOf } from '../opc/reproducible';
import { freshControlId, type ContentControl } from './content-control-model';
import type { Block } from './model';
import { isSdt, sdtContent } from './content-control';
import {
	createContentControl,
	patchContentControl,
	unwrapContentControl,
} from './write-content-control';
import type { FlatBlock } from './parse-content-control';
import { replaceSlots } from './write-slots';
import { first, getW, isElement, named, WORD_NS, type XmlDocument, type XmlElement } from './xml';

/** A block or row as the placement reads its controls. */
export type ControlHolder = Pick<Block, 'contentControl' | 'outerContentControls'> | undefined;

/** The controls around a block, outermost first, ending with its own; none without one. */
export const blockChain = (block: ControlHolder): ContentControl[] =>
	block?.contentControl ? [...(block.outerContentControls ?? []), block.contentControl] : [];

/**
 * Whether the model says which controls are around the block: `outerContentControls` set, even
 * empty. A model that leaves it unset (an editor that carries only the block's own control)
 * leaves the `w:sdt` elements around that control as they are.
 */
const tracksOuter = (block: ControlHolder): boolean => block?.outerContentControls !== undefined;

/** The model id of each source `w:sdt` (a control read without `w:id` has a generated one). */
function sourceIds(flat: readonly FlatBlock[]): Map<XmlElement, string> {
	const ids = new Map<XmlElement, string>();
	for (const entry of flat) {
		if (entry.sdt && entry.control) ids.set(entry.sdt, entry.control.id);
		for (const level of entry.outer ?? []) ids.set(level.sdt, level.control.id);
	}
	return ids;
}

/** The control ids every `w:sdtPr` of `doc` has. */
function partControlIds(doc: XmlDocument): Set<string> {
	const used = new Set<string>();
	for (const id of Array.from(doc.getElementsByTagNameNS(WORD_NS, 'id')))
		if (named(id.parentNode as XmlElement, 'sdtPr')) used.add(getW(id, 'val') ?? '');
	return used;
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

/**
 * Where a new element goes beside `neighbour` (a placed one): inside the `w:sdt` elements around
 * `neighbour` that `chain` shares with it (by model id, outermost first), or every one up to the
 * level holding `chain`'s own control when `chain` names only some of the controls around it.
 * Returns the number of shared levels and the controls still to wrap it in.
 */
function sharedLevels(
	sdts: readonly XmlElement[],
	chain: readonly ContentControl[],
	idOf: (sdt: XmlElement) => string | undefined,
): { depth: number; rest: readonly ContentControl[] } {
	const ids = sdts.map(idOf);
	const own = chain.at(-1);
	const at = own ? ids.lastIndexOf(own.id) : -1;
	if (at >= 0) return { depth: at + 1, rest: [] };
	let depth = 0;
	while (depth < Math.min(chain.length, ids.length) && ids[depth] === chain[depth]!.id) depth++;
	return { depth, rest: chain.slice(depth) };
}

/**
 * Places elements already written to `nodes` in `container`, each inside the controls of
 * `holders` (parallel to `nodes`). Returns false, changing nothing, when the kept elements
 * changed their relative order (the caller then replaces slot by slot).
 */
export function placeControlledById(
	doc: XmlDocument,
	container: XmlElement,
	flat: FlatBlock[],
	holders: readonly ControlHolder[],
	nodes: XmlElement[],
	anchor: XmlElement | null,
): boolean {
	const chains = holders.map(blockChain);
	const modelIds = sourceIds(flat);
	const idOf = (sdt: XmlElement) =>
		modelIds.get(sdt) ?? getW(first(first(sdt, 'sdtPr'), 'id'), 'val');
	let used: Set<string> | undefined;
	/** Wraps `node` where it is in a new `w:sdt` for each control, the first outermost. */
	const wrap = (node: XmlElement, controls: readonly ContentControl[]) => {
		const around = new Set(sdtAncestors(node, container).map(idOf));
		for (const control of controls) {
			// Already inside it: never a second `w:sdt` with its id around the same content.
			if (around.has(control.id)) continue;
			used ??= partControlIds(doc);
			// Ids are unique in a part (ECMA-376 17.5.2.18): one in use elsewhere gets a new one.
			const id = used.has(control.id)
				? freshControlId(used, saveEnvironmentOf(doc).ids)
				: control.id;
			used.add(id);
			const { sdt, content } = createContentControl(
				doc,
				id === control.id ? control : { ...control, id },
			);
			modelIds.set(sdt, control.id);
			node.parentNode!.insertBefore(sdt, node);
			content.appendChild(node);
		}
	};
	const oldSet = new Set(flat.map((entry) => entry.element));
	const outputSet = new Set(nodes);
	const keptInOrder = nodes.filter((node) => oldSet.has(node));
	const documentOrder = flat.map((entry) => entry.element).filter((node) => outputSet.has(node));
	if (keptInOrder.some((node, index) => node !== documentOrder[index])) return false;
	for (const entry of flat)
		if (!outputSet.has(entry.element)) entry.element.parentNode?.removeChild(entry.element);
	pruneEmptyControls(flat);
	let previous: XmlElement | undefined;
	nodes.forEach((node, index) => {
		const chain = chains[index] ?? [];
		if (oldSet.has(node)) {
			previous = node;
			return;
		}
		// After the element before it, else before the first kept one, inside what they share.
		const neighbour = previous ?? keptInOrder[0];
		if (neighbour) {
			const sdts = sdtAncestors(neighbour, container);
			const { depth, rest } = sharedLevels(sdts, chain, idOf);
			const reference = depth >= sdts.length ? neighbour : sdts[depth]!;
			const parent = reference.parentNode as XmlElement;
			parent.insertBefore(node, previous ? reference.nextSibling : reference);
			wrap(node, rest);
		} else {
			container.insertBefore(node, anchor);
			wrap(node, chain);
		}
		previous = node;
	});
	syncControls(doc, flat, holders, nodes);
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
): boolean => placeControlledById(doc, container, flat, blocks, nodes, anchor);

/**
 * Brings the `w:sdt` wrappers of `flat` (the source elements, before the edit) in step with the
 * model: edited control properties reach the `w:sdt` that already wraps an element, and a control
 * the model took off every element it held is unwrapped (a temporary control after an edit, a
 * flattened one). `nodes` are the written elements, parallel to `chains`.
 */
export function syncControls(
	doc: XmlDocument,
	flat: FlatBlock[],
	holders: readonly ControlHolder[],
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
		// A deleted element has no say: the ones that remain decide whether their control stays.
		if (!entry.sdt || index === undefined) return;
		const holder = holders[index];
		// A model that does not say which controls are around the element's own follows only the
		// innermost `w:sdt`: patched from its control, unwrapped when it has none.
		if (!tracksOuter(holder)) {
			note(entry.sdt, holder?.contentControl);
			return;
		}
		// Each `w:sdt` around it follows the model's control with its id: patched when the model
		// keeps it, unwrapped when the model took it off every element it held. A model that gives
		// the innermost one another id (and no outer one) patches it, as one control.
		const chain = blockChain(holder);
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
): void => syncControls(doc, flat, blocks, nodes);

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
