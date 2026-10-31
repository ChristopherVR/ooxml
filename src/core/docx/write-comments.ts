// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { Comment } from './model';
import {
	buildXml,
	getW,
	makeW,
	parseXml,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
	WORD_DATE_UTC_NS,
} from './xml';
import { ensureContentTypeOverride, ensureDocumentRelationship } from './zip-parts';
import { COMMENT_PARTS, readCommentParts, W14_NS, type CommentPartKind } from './comment-parts';
import { commentParagraphId, parseComments } from './comments';
import { commentHexIds, commentNumericIds, patchCommentAnchors } from './write-comment-identities';
import {
	existingDurableIds,
	patchCommentMetadata,
	type CommentIdentity,
} from './write-comment-metadata';
import { patchCommentText } from './write-comment-text';
import { declareIgnorableExtension } from './ligatures';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

/** Edit comment bodies and metadata in their source parts, retaining unchanged formatting and IDs. */
export async function applyComments(zip: JSZip, comments: Comment[]): Promise<void> {
	const parts = await readCommentParts(zip);
	const previous = parts.comments.xml
		? parseComments(parts.comments.xml, parts.extended.xml, parts.ids.xml, parts.extensible.xml)
		: [];
	const byId = new Map(previous.map((comment) => [comment.id, comment]));
	const ids = commentNumericIds(comments, previous);
	const nextHexId = await commentHexIds(zip);
	const durableIds = existingDurableIds(parts);
	const doc = parseXml(
		parts.comments.xml ??
			`<w:comments xmlns:w="${WORD_NS}" xmlns:w14="${W14_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"/>`,
	);
	const before = buildXml(doc);
	const w14 = declareIgnorableExtension(doc, W14_NS, 'w14');
	const elements = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'comment')) as XmlElement[];
	const sourceElements = new Map(elements.map((element) => [getW(element, 'id'), element]));
	const retained = new Set<XmlElement>();
	const identities: CommentIdentity[] = [];
	for (const comment of comments) {
		const old = byId.get(comment.id);
		const element = sourceElements.get(comment.id) ?? makeW(doc, 'comment');
		const paraId = commentParagraphId(element) ?? nextHexId();
		setAttribute(element, 'id', ids.get(comment.id)!);
		for (const key of ['author', 'initials', 'date'] as const) {
			if (old && comment[key] === old[key]) continue;
			const value = comment[key];
			if (value === undefined) element.removeAttributeNS(WORD_NS, key);
			else setAttribute(element, key, value);
			if (key === 'date') {
				const utc = element.getAttributeNodeNS(WORD_DATE_UTC_NS, 'dateUtc');
				if (utc && value === undefined) element.removeAttributeNode(utc);
				else if (utc) utc.value = value!;
			}
		}
		if (!old || old.text !== comment.text) patchCommentText(element, comment.text);
		const paragraphs = Array.from(element.getElementsByTagNameNS(WORD_NS, 'p'));
		if (!paragraphs.length) {
			const paragraph = makeW(doc, 'p');
			element.appendChild(paragraph);
			paragraphs.push(paragraph);
		}
		for (const paragraph of paragraphs) {
			const oldId = paragraph.getAttributeNS(W14_NS, 'paraId');
			if (paragraph === paragraphs.at(-1))
				paragraph.setAttributeNS(W14_NS, `${w14}:paraId`, paraId);
			else if (!oldId || oldId.toUpperCase() === paraId.toUpperCase())
				paragraph.setAttributeNS(W14_NS, `${w14}:paraId`, nextHexId());
		}
		if (!element.parentNode) doc.documentElement!.appendChild(element);
		retained.add(element);
		identities.push({
			comment,
			previous: old,
			paraId,
			durableId: durableIds.get(paraId.toUpperCase()) ?? nextHexId(),
		});
	}
	for (const element of elements)
		if (!retained.has(element)) element.parentNode?.removeChild(element);
	const after = buildXml(doc);
	zip.file(
		parts.comments.path,
		parts.comments.xml && before === after ? parts.comments.xml : after,
	);
	const metadata = patchCommentMetadata(parts, identities);
	for (const [path, xml] of metadata) zip.file(path, xml);
	for (const kind of Object.keys(COMMENT_PARTS) as CommentPartKind[]) {
		const part = parts[kind];
		if (kind !== 'comments' && !metadata.has(part.path)) continue;
		await ensureContentTypeOverride(
			zip,
			part.path,
			`application/vnd.openxmlformats-officedocument.wordprocessingml.${COMMENT_PARTS[kind].name}+xml`,
		);
		if (!part.related)
			await ensureDocumentRelationship(zip, COMMENT_PARTS[kind].relationship, `/${part.path}`);
	}
	await patchCommentAnchors(zip, ids, comments);
}

export function commentRangeStartNode(doc: XmlDocument, id: string): XmlElement {
	const element = makeW(doc, 'commentRangeStart');
	setAttribute(element, 'id', id);
	return element;
}
/** `commentRangeEnd` plus the anchoring `commentReference` run that follows it. */
export function commentRangeEndNodes(doc: XmlDocument, id: string): XmlElement[] {
	const end = makeW(doc, 'commentRangeEnd');
	setAttribute(end, 'id', id);
	const run = makeW(doc, 'r');
	const reference = makeW(doc, 'commentReference');
	setAttribute(reference, 'id', id);
	run.appendChild(reference);
	return [end, run];
}
