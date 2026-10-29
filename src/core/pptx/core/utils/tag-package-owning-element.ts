/**
 * Author/remove the `<p:tags r:id=".."/>` (CT_TagsData) element that OWNS a
 * tags-part relationship.
 *
 * `p:tags` is a child of `p:custDataLst` (CT_CustomerDataList), itself a
 * child of either `p:presentation` or a `p:cSld`-bearing part (slide, notes
 * slide, slide layout, slide master, handout master - CT_CommonSlideData is
 * shared by all of them), or of a shape's `p:nvPr`. Without this element,
 * real PowerPoint's smart-tags UI never sees a tags part that only exists as a
 * bare relationship - see `src/__tests__/integration/tag-part-authoring.test.ts`
 * for the authoring contract this closes.
 *
 * These passes run AFTER the slide writer, on parts already serialized into
 * the zip, so they edit the part through the shared `xml` DOM rather than the
 * fast-xml-parser object model. That model groups same-tag siblings into one
 * array per tag, and its builder never consults the parse-time order
 * annotations, so a parse/build round trip here rewrote every interleaved
 * sequence in the slide grouped by tag: custom-geometry path commands came
 * out as all `moveTo`s, then all `lnTo`s (a freeform icon saved as a tangle),
 * the shape tree lost its z-order and inline fields moved to the end of their
 * paragraphs. The DOM keeps document order, and a part whose owning element
 * is already present is not rewritten at all.
 *
 * @module tag-package-owning-element
 */
import type JSZip from 'jszip';

import {
	NS,
	buildXml,
	parseXml,
	relAttr,
	type XmlDocument,
	type XmlElement,
} from '../../../xml/index';
import {
	createPresentationElement,
	custDataContainer,
	ensureCustDataList,
	listChildren,
	removeIfEmpty,
} from './cust-data-list-dom';

/** Every `p:tags` element in the part (part-level and shape-level alike). */
function tagsElements(doc: XmlDocument): XmlElement[] {
	return Array.from(doc.getElementsByTagNameNS(NS.p, 'tags')) as XmlElement[];
}

/** Whether `xml` may carry a `p:tags` whose relationship id is `relationshipId`. */
function mentionsTagsId(xml: string, relationshipId: string): boolean {
	return xml.includes('tags') && xml.includes(`"${relationshipId}"`);
}

/**
 * Write `<p:tags r:id="..">` into the owning part's `p:custDataLst`, merging
 * with any existing `p:custData` entries and preserving unknown XML.
 *
 * A part that already carries a `p:tags` for `relationshipId` (wherever it
 * sits: the part's own list or a shape's `p:nvPr`) is left byte-for-byte as
 * it is. Re-placing it would move a shape's tags onto the slide and, with
 * several collections, leave the slide pointing at the last shape's part.
 */
export async function upsertTagsOwningElement(
	zip: JSZip,
	sourcePartPath: string,
	relationshipId: string,
): Promise<void> {
	const xml = await zip.file(sourcePartPath)?.async('string');
	if (!xml) {
		return;
	}
	const doc = parseXml(xml, { label: sourcePartPath });
	if (
		mentionsTagsId(xml, relationshipId) &&
		tagsElements(doc).some((tags) => relAttr(tags, 'id') === relationshipId)
	) {
		return;
	}
	const container = custDataContainer(doc);
	if (!container) {
		return;
	}
	const list = ensureCustDataList(doc, container);
	// CT_CustomerDataList is `custData*, tags?`: one list owns at most one tags part.
	for (const stale of listChildren(list, 'tags')) {
		list.removeChild(stale);
	}
	const tags = createPresentationElement(doc, container, 'tags');
	tags.setAttributeNS(NS.r, 'r:id', relationshipId);
	list.appendChild(tags);
	zip.file(sourcePartPath, buildXml(doc));
}

/**
 * Remove the `<p:tags r:id=".."/>` that owns `relationshipId` from the part,
 * wherever it sits, dropping its `p:custDataLst` once that carries neither
 * tags nor custom data. Without a relationship id, the part's own list loses
 * its `p:tags`.
 */
export async function removeTagsOwningElement(
	zip: JSZip,
	sourcePartPath: string,
	relationshipId: string | undefined,
): Promise<void> {
	const xml = await zip.file(sourcePartPath)?.async('string');
	if (!xml || (relationshipId !== undefined && !mentionsTagsId(xml, relationshipId))) {
		return;
	}
	const doc = parseXml(xml, { label: sourcePartPath });
	const container = custDataContainer(doc);
	const owned = tagsElements(doc).filter((tags) =>
		relationshipId === undefined
			? tags.parentNode?.parentNode === container
			: relAttr(tags, 'id') === relationshipId,
	);
	if (owned.length === 0) {
		return;
	}
	for (const tags of owned) {
		const list = tags.parentNode as XmlElement | null;
		list?.removeChild(tags);
		if (list) {
			removeIfEmpty(list);
		}
	}
	zip.file(sourcePartPath, buildXml(doc));
}
