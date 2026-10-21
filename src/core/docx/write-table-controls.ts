// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Keeps the content controls around table rows and cells in step with the model: an edited control
// is patched in place, a control the model dropped from every row or cell it holds is unwrapped
// (they stay), and a control the model added wraps its row or cell in a new `w:sdt`.
import type { ContentControl } from './content-control-model';
import { isSdt, sdtContent } from './content-control';
import {
	createContentControl,
	patchContentControl,
	unwrapContentControl,
} from './write-content-control';
import { freshControlId } from './content-control-model';
import { first, getW, named, WORD_NS, type XmlDocument, type XmlElement } from './xml';

/** The `w:sdt` whose content directly holds `element`, if any. */
function enclosingSdt(element: XmlElement): XmlElement | undefined {
	const content = element.parentNode as XmlElement | null;
	const sdt = content?.parentNode as XmlElement | null | undefined;
	return named(content, 'sdtContent') && isSdt(sdt) ? sdt : undefined;
}

/** Copies of a `w:sdt` split by a row or cell given another control, under new ids. */
const renumbered = new WeakSet<XmlElement>();
/**
 * The model's id of each `w:sdt` written under a new id because another `w:sdt` has the model's
 * (a new control's, or the rest of a split one), or under none because Word cannot read it: the
 * rows or cells after it with that id join it.
 */
const createdFor = new WeakMap<XmlElement, string>();
const idOf = (sdt: XmlElement): string | undefined => getW(first(first(sdt, 'sdtPr'), 'id'), 'val');

/** The ids of the controls in `doc`. */
function controlIds(doc: XmlDocument): Set<string> {
	const used = new Set<string>();
	for (const id of Array.from(doc.getElementsByTagNameNS(WORD_NS, 'id')))
		if (named(id.parentNode as XmlElement, 'sdtPr')) used.add(getW(id, 'val') ?? '');
	return used;
}

/** A control id that no `w:sdtPr` in `doc` has. */
const unusedControlId = (doc: XmlDocument): string => freshControlId(controlIds(doc));

/**
 * Whether `node` is a row or cell, or holds rows or cells (a `w:sdt` or `w:customXml` around them):
 * what a row or cell control holds beside range markers such as bookmarks, comment and permission
 * ranges and proofing marks (EG_ContentRowContent, EG_ContentCellContent).
 */
const holdsRowsOrCells = (node: ChildNode | null): boolean =>
	node?.nodeType === 1 &&
	['tr', 'tc', 'sdt', 'customXml'].some((name) => named(node as XmlElement, name));

/**
 * Moves the range markers Word can write beside a row or cell control's `w:sdtContent` (MS-OI29500
 * 2.1.202 and 2.1.204) into it, where they are in reading order: before what it holds, or after.
 * Its content allows them (EG_ContentRowContent, EG_ContentCellContent), so a row or cell taken out
 * of the control, or joining it, keeps them around what they held.
 */
function markersIntoContent(sdt: XmlElement): void {
	const content = sdtContent(sdt);
	if (!content) return;
	const anchor = content.firstChild;
	let after = false;
	for (const child of Array.from(sdt.childNodes)) {
		if (child === content) {
			after = true;
			continue;
		}
		const element = child as XmlElement;
		if (child.nodeType !== 1 || named(element, 'sdtPr') || named(element, 'sdtEndPr')) continue;
		if (after) content.appendChild(child);
		else content.insertBefore(child, anchor);
	}
}

/**
 * Takes `element` out of `sdt`, which is another control's: at its start or end it moves out beside
 * it, with the range markers on that side, and from its middle it goes after it, the rows or cells
 * after it going to a copy of `sdt` under a new id (ids are unique, ECMA-376 17.5.2.18). Everything
 * keeps its order. An `sdt` left without rows or cells is unwrapped.
 */
function takeOut(doc: XmlDocument, element: XmlElement, sdt: XmlElement): void {
	markersIntoContent(sdt);
	const content = element.parentNode as XmlElement;
	const parent = sdt.parentNode as XmlElement;
	const siblings = (from: ChildNode | null, step: 'previousSibling' | 'nextSibling') => {
		const found: ChildNode[] = [];
		for (let node = from; node; node = node[step]) found.push(node);
		return found;
	};
	const before = siblings(element.previousSibling, 'previousSibling');
	const after = siblings(element.nextSibling, 'nextSibling');
	if (!before.some(holdsRowsOrCells)) {
		for (const node of before.reverse()) parent.insertBefore(node, sdt);
		parent.insertBefore(element, sdt);
	} else if (!after.some(holdsRowsOrCells)) {
		const next = sdt.nextSibling;
		for (const node of [element, ...after]) parent.insertBefore(node, next);
	} else {
		const tail = sdt.cloneNode(false) as XmlElement;
		for (const child of Array.from(sdt.childNodes))
			if (named(child as XmlElement, 'sdtPr') || named(child as XmlElement, 'sdtEndPr'))
				tail.appendChild(child.cloneNode(true));
		const id = first(first(tail, 'sdtPr'), 'id');
		const model = createdFor.get(sdt) ?? idOf(sdt);
		if (id) id.setAttributeNS(WORD_NS, 'w:val', unusedControlId(doc));
		renumbered.add(tail);
		if (model !== undefined) createdFor.set(tail, model);
		const rest = content.cloneNode(false) as XmlElement;
		for (const node of after) rest.appendChild(node);
		tail.appendChild(rest);
		parent.insertBefore(tail, sdt.nextSibling);
		parent.insertBefore(element, tail);
	}
	if (!Array.from(content.childNodes).some(holdsRowsOrCells)) unwrapContentControl(sdt);
}

/**
 * Writes the control of the row or cell `element` (already in place in its table): `next` is the
 * model's control, `base` the one it was loaded with. An unchanged control is left as it was; a
 * dropped one is unwrapped by `unwrapDroppedControls`, which sees every row or cell it holds.
 */
export function placeTableControl(
	doc: XmlDocument,
	element: XmlElement,
	next: ContentControl | undefined,
	base: ContentControl | undefined,
): void {
	if (!next || JSON.stringify(next) === JSON.stringify(base)) return;
	const sdt = enclosingSdt(element);
	// The `w:sdt` around it is this control only if it was loaded there (the same id; patching
	// replaces its type element if its kind changed): given another control, the row or cell leaves
	// it for a `w:sdt` of its own, built from the model. A copy split off keeps its own id.
	if (sdt && base?.id === next.id) {
		patchContentControl(
			doc,
			sdt,
			renumbered.has(sdt) ? { ...next, id: idOf(sdt) ?? next.id } : next,
		);
		return;
	}
	if (sdt) takeOut(doc, element, sdt);
	if (!element.parentNode) return;
	// A row or cell given the control of the one before it joins that one's `w:sdt`, also when it
	// was written under a new id, with the range markers between them.
	const between: ChildNode[] = [];
	let before = element.previousSibling;
	while (before && !holdsRowsOrCells(before)) {
		between.unshift(before);
		before = before.previousSibling;
	}
	const shared = isSdt(before as XmlElement | null) && sdtContent(before as XmlElement);
	const { id } = next;
	const joins =
		shared &&
		typeof id === 'string' &&
		(createdFor.get(before as XmlElement) ?? idOf(before as XmlElement)) === id;
	// A control without an id of its own (a remote peer can give one) joins none.
	if (joins) {
		markersIntoContent(before as XmlElement);
		for (const node of [...between, element]) shared.appendChild(node);
		return;
	}
	// Ids are unique in a document: a control given the id of another `w:sdt` gets a new one.
	const taken = controlIds(doc).has(id);
	const fresh = createContentControl(doc, taken ? { ...next, id: unusedControlId(doc) } : next);
	// Written under another id, or none when Word cannot read the model's, it is still the one the
	// rows or cells after it with the model's id join.
	if (idOf(fresh.sdt) !== id) createdFor.set(fresh.sdt, id);
	element.parentNode.replaceChild(fresh.sdt, element);
	fresh.content.appendChild(element);
}

/**
 * Unwraps each `w:sdt` around `elements` (the rows of a table or the cells of a row, in place)
 * that the model takes off every one of them it holds: `next` is each element's control. The
 * rows or cells stay; a control the model keeps on any of them stays around all of them.
 */
export function unwrapDroppedControls(
	elements: readonly XmlElement[],
	next: readonly (ContentControl | undefined)[],
): void {
	const kept = new Map<XmlElement, boolean>();
	elements.forEach((element, index) => {
		const sdt = enclosingSdt(element);
		if (sdt) kept.set(sdt, Boolean(kept.get(sdt)) || next[index] !== undefined);
	});
	for (const [sdt, keep] of kept) if (!keep) unwrapContentControl(sdt);
}
