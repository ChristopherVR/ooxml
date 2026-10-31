import type { CommentParts } from './comment-parts';
import { W15_NS, W16CID_NS, W16CEX_NS } from './comment-parts';
import type { Comment } from './model';
import { buildXml, parseXml, type XmlElement } from './xml';

export interface CommentIdentity {
	comment: Comment;
	previous: Comment | undefined;
	paraId: string;
	durableId: string;
}

export function existingDurableIds(parts: CommentParts): Map<string, string> {
	const ids = new Map<string, string>();
	if (parts.ids.xml)
		for (const entry of Array.from(
			parseXml(parts.ids.xml).getElementsByTagNameNS(W16CID_NS, 'commentId'),
		)) {
			const paragraph = entry.getAttributeNS(W16CID_NS, 'paraId');
			const durable = entry.getAttributeNS(W16CID_NS, 'durableId');
			if (paragraph && durable) ids.set(paragraph.toUpperCase(), durable);
		}
	return ids;
}

/** Update known attributes on matching entries while retaining extension markup and root metadata. */
export function patchCommentMetadata(
	parts: CommentParts,
	identities: CommentIdentity[],
): Map<string, string> {
	const result = new Map<string, string>();
	const paragraphs = new Map(identities.map((entry) => [entry.comment.id, entry.paraId]));
	const definitions = [
		{
			kind: 'extended',
			ns: W15_NS,
			prefix: 'w15',
			root: 'commentsEx',
			entry: 'commentEx',
			key: 'paraId',
		},
		{
			kind: 'ids',
			ns: W16CID_NS,
			prefix: 'w16cid',
			root: 'commentsIds',
			entry: 'commentId',
			key: 'paraId',
		},
		{
			kind: 'extensible',
			ns: W16CEX_NS,
			prefix: 'w16cex',
			root: 'commentsExtensible',
			entry: 'commentExtensible',
			key: 'durableId',
		},
	] as const;
	for (const definition of definitions) {
		const { kind, ns, prefix, root, entry, key } = definition;
		const source = parts[kind].xml;
		if (!source && !identities.length) continue;
		const doc = parseXml(source ?? `<${prefix}:${root} xmlns:${prefix}="${ns}"/>`);
		const before = buildXml(doc);
		const elements = Array.from(doc.getElementsByTagNameNS(ns, entry)) as XmlElement[];
		const byKey = new Map(
			elements.map((element) => [element.getAttributeNS(ns, key)?.toUpperCase(), element]),
		);
		const retained = new Set<XmlElement>();
		for (const identity of identities) {
			const { comment, previous } = identity;
			const id = key === 'paraId' ? identity.paraId : identity.durableId;
			const element = byKey.get(id.toUpperCase()) ?? doc.createElementNS(ns, `${prefix}:${entry}`);
			const set = (name: string, value: string | undefined) => {
				if (value === undefined) element.removeAttributeNS(ns, name);
				else if (element.getAttributeNS(ns, name) !== value)
					element.setAttributeNS(ns, `${prefix}:${name}`, value);
			};
			set(key, id);
			if (kind === 'extended') {
				if (!previous || comment.resolved !== previous.resolved)
					set('done', comment.resolved ? '1' : '0');
				set('paraIdParent', comment.parentId ? paragraphs.get(comment.parentId) : undefined);
			} else if (kind === 'ids') set('durableId', identity.durableId);
			else if (
				!previous ||
				comment.date !== previous.date ||
				!element.hasAttributeNS(ns, 'dateUtc')
			)
				set('dateUtc', comment.date);
			if (!element.parentNode) doc.documentElement!.appendChild(element);
			retained.add(element);
		}
		for (const element of elements)
			if (!retained.has(element)) element.parentNode?.removeChild(element);
		const after = buildXml(doc);
		result.set(parts[kind].path, source && before === after ? source : after);
	}
	return result;
}
