// Writes editable ranges (`w:permStart`/`w:permEnd`) the Restrict Editing pane changed. A paragraph
// whose `permissions` still equal the source keeps its markers as they were (inline-markers.ts
// carries them through edits); a changed one gets its markers rewritten from the model.
import type { Paragraph } from './model';
import { makeW, named, WORD_NS, type XmlDocument, type XmlElement } from './xml';

const setAttribute = (element: XmlElement, local: string, value: string) =>
	element.setAttributeNS(WORD_NS, `w:${local}`, value);

const isPermission = (element: XmlElement) =>
	named(element, 'permStart') || named(element, 'permEnd');

/** A written marker and the character position it goes to. */
export interface PermissionMarker {
	element: XmlElement;
	position: number;
	source: number;
	opening: boolean;
}

/**
 * The markers to drop from the source paragraph and the ones to add, or undefined when the
 * paragraph's ranges are as they were (or it is new: its ranges cannot be told from the stale
 * copy a split paragraph carries).
 */
export function changedPermissionMarkers(
	doc: XmlDocument,
	paragraph: Paragraph,
	base: Paragraph | undefined,
	source: readonly XmlElement[],
	length: number,
): { drop: Set<XmlElement>; added: PermissionMarker[] } | undefined {
	if (
		!base ||
		JSON.stringify(paragraph.permissions ?? null) === JSON.stringify(base.permissions ?? null)
	)
		return undefined;
	// A range read from a file that goes on into another paragraph keeps its markers where they are.
	const spans = (list: Paragraph['permissions']) => list?.some((r) => r.fromPrevious || r.toNext);
	if (spans(paragraph.permissions) || spans(base.permissions)) return undefined;
	const drop = new Set(source.filter(isPermission));
	const added: PermissionMarker[] = [];
	for (const range of paragraph.permissions ?? []) {
		const start = makeW(doc, 'permStart');
		setAttribute(start, 'id', range.id);
		if (range.group !== undefined) setAttribute(start, 'edGrp', range.group);
		if (range.user !== undefined) setAttribute(start, 'ed', range.user);
		const end = makeW(doc, 'permEnd');
		setAttribute(end, 'id', range.id);
		added.push(
			{ element: start, position: Math.min(range.from, length), source: range.from, opening: true },
			{ element: end, position: Math.min(range.to, length), source: range.to, opening: false },
		);
	}
	return { drop, added };
}
