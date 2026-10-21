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
 * (a new control's, or the rest of a split one): the rows or cells after it with that id join it.
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
 * Takes `element` out of `sdt`, which is another control's: at its start or end it moves out beside
 * it, and from its middle it goes after it, the rows or cells after it going to a copy of `sdt`
 * under a new id (ids are unique, ECMA-376 17.5.2.18). An `sdt` left empty is unwrapped.
 */
function takeOut(doc: XmlDocument, element: XmlElement, sdt: XmlElement): void {
	const content = element.parentNode as XmlElement;
	const parent = sdt.parentNode as XmlElement;
	const isElement = (node: ChildNode | null) => node?.nodeType === 1;
	const siblings = (from: ChildNode | null, step: 'previousSibling' | 'nextSibling') => {
		const found: ChildNode[] = [];
		for (let node = from; node; node = node[step]) found.push(node);
		return found;
	};
	const before = siblings(element.previousSibling, 'previousSibling');
	const after = siblings(element.nextSibling, 'nextSibling');
	if (!before.some(isElement)) parent.insertBefore(element, sdt);
	else if (!after.some(isElement)) parent.insertBefore(element, sdt.nextSibling);
	else {
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
	if (!Array.from(content.childNodes).some(isElement)) unwrapContentControl(sdt);
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
	// was written under a new id.
	let before = element.previousSibling;
	while (before && before.nodeType !== 1) before = before.previousSibling;
	const shared = isSdt(before as XmlElement | null) && sdtContent(before as XmlElement);
	const id = /^-?\d+$/.test(next.id) ? next.id : undefined;
	if (
		shared &&
		id !== undefined &&
		(createdFor.get(before as XmlElement) ?? idOf(before as XmlElement)) === id
	) {
		shared.appendChild(element);
		return;
	}
	// Ids are unique in a document: a control given the id of another `w:sdt` gets a new one.
	const taken = id !== undefined && controlIds(doc).has(id);
	const fresh = createContentControl(doc, taken ? { ...next, id: unusedControlId(doc) } : next);
	if (taken) createdFor.set(fresh.sdt, id);
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
