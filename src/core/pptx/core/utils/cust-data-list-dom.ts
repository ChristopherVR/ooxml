/**
 * The part-level `p:custDataLst` (CT_CustomerDataList) of an already-written
 * package part, edited through the shared `xml` DOM.
 *
 * The tag and customer-data passes run after the slide writer, on parts that
 * are already serialized. Re-parsing such a part with fast-xml-parser and
 * building it again regroups every interleaved run of siblings by tag (custom
 * geometry path commands, the shape tree's z-order, inline fields), because the
 * builder ignores the parse-time order annotations. The DOM keeps document
 * order, so these passes touch only the list they own.
 *
 * @module cust-data-list-dom
 */
import { NS, elements, first, type XmlDocument, type XmlElement } from '../../../xml/index';

/** Package parts whose content sits under a `p:cSld` (CT_CommonSlideData). */
const CSLD_ROOTS = new Set(['sld', 'notes', 'sldLayout', 'sldMaster', 'handoutMaster']);

/** CT_Presentation child order: ... custDataLst, kinsoku, defaultTextStyle, modifyVerifier, extLst. */
const PRESENTATION_INSERT_BEFORE = new Set([
	'kinsoku',
	'defaultTextStyle',
	'modifyVerifier',
	'extLst',
]);
/** CT_CommonSlideData child order: bg, spTree, custDataLst, controls, extLst. */
const CSLD_INSERT_BEFORE = new Set(['controls', 'extLst']);

/** The element whose own `p:custDataLst` belongs to the part: `p:presentation` or `p:cSld`. */
export function custDataContainer(doc: XmlDocument): XmlElement | undefined {
	const root = doc.documentElement;
	if (!root || root.namespaceURI !== NS.p) {
		return undefined;
	}
	if (root.localName === 'presentation') {
		return root;
	}
	return CSLD_ROOTS.has(root.localName) ? first(root, 'cSld', NS.p) : undefined;
}

/** A `p:` element named `local`, spelled with `container`'s own prefix. */
export function createPresentationElement(
	doc: XmlDocument,
	container: XmlElement,
	local: string,
): XmlElement {
	return doc.createElementNS(NS.p, container.prefix ? `${container.prefix}:${local}` : local);
}

/** The container's `p:custDataLst`, created in its schema position when absent. */
export function ensureCustDataList(doc: XmlDocument, container: XmlElement): XmlElement {
	const existing = first(container, 'custDataLst', NS.p);
	if (existing) {
		return existing;
	}
	const insertBefore =
		container.localName === 'presentation' ? PRESENTATION_INSERT_BEFORE : CSLD_INSERT_BEFORE;
	const list = createPresentationElement(doc, container, 'custDataLst');
	const before = elements(container).find(
		(child) => child.namespaceURI === NS.p && insertBefore.has(child.localName),
	);
	container.insertBefore(list, before ?? null);
	return list;
}

/** The direct `p:<local>` children of a list. */
export function listChildren(list: XmlElement, local: string): XmlElement[] {
	return elements(list).filter((child) => child.namespaceURI === NS.p && child.localName === local);
}

/** Drop `list` once it has no element children left. */
export function removeIfEmpty(list: XmlElement): void {
	if (elements(list).length === 0) {
		list.parentNode?.removeChild(list);
	}
}
