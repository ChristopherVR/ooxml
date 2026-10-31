import { NS, buildXml, children, first, parseXml } from '../../xml/index';
import { parseAddress } from '../address';
import type { Comment, CommentMessage } from '../model';
import { att, boolAttr, selfContainedXml, xChildren, xFirst, xText } from './xml-util';

/** Legacy comments (notes): the author list and each comment's plain text. */
export function parseLegacyComments(xml: string | undefined): Comment[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'XLSX comments' }).documentElement;
	const authors = xChildren(xFirst(root, 'authors') ?? root, 'author').map((a) => xText(a));
	const out: Comment[] = [];
	for (const node of xChildren(xFirst(root, 'commentList') ?? root, 'comment')) {
		const address = parseAddress(att(node, 'ref') ?? '');
		if (!address) continue;
		const text = xFirst(node, 'text');
		const runs = text ? xChildren(text, 'r') : [];
		const body = runs.length
			? runs.map((run) => xText(xFirst(run, 't'))).join('')
			: xText(xFirst(text, 't'));
		out.push({ address, author: authors[Number(att(node, 'authorId') ?? 0)] ?? '', text: body });
	}
	return out;
}

/** `person id -> display name` from `xl/persons/person.xml`. */
export function parsePersons(xml: string | undefined): Map<string, string> {
	const persons = new Map<string, string>();
	if (!xml) return persons;
	const root = parseXml(xml, { label: 'XLSX persons' }).documentElement;
	for (const person of children(root, 'person', NS.tc)) {
		const id = att(person, 'id');
		if (id) persons.set(id, att(person, 'displayName') ?? '');
	}
	return persons;
}

/** Threaded comments grouped into root comments with their replies (oldest first). */
export function parseThreadedComments(
	xml: string | undefined,
	persons: ReadonlyMap<string, string>,
	personsXml?: string,
): Comment[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'XLSX threaded comments' }).documentElement;
	const roots = new Map<string, Comment>();
	const personNames = new Map([...persons].map(([id, name]) => [id.toUpperCase(), name]));
	const order: Comment[] = [];
	const nodes = children(root, 'threadedComment', NS.tc);
	const byId = new Map(nodes.map((node) => [(att(node, 'id') ?? '').toUpperCase(), node]));
	const parentRoot = (parent: string | undefined): Comment | undefined => {
		const seen = new Set<string>();
		while (parent) {
			parent = parent.toUpperCase();
			if (seen.has(parent)) break;
			if (roots.has(parent)) return roots.get(parent);
			seen.add(parent);
			parent = att(byId.get(parent), 'parentId');
		}
		return undefined;
	};
	// References need not precede their users in the part, including reply-to-reply references.
	for (const node of [
		...nodes.filter((n) => !att(n, 'parentId')),
		...nodes.filter((n) => att(n, 'parentId')),
	]) {
		const parentComment = parentRoot(att(node, 'parentId'));
		const address = parseAddress(att(node, 'ref') ?? '') ?? parentComment?.address;
		if (!address) continue;
		const author = personNames.get((att(node, 'personId') ?? '').toUpperCase()) ?? '';
		const text = xText(children(node, 'text', NS.tc)[0]);
		const date = att(node, 'dT');
		if (parentComment) {
			const reply: CommentMessage = { author, text, threadedXml: selfContainedXml(node) };
			if (date) reply.date = date;
			(parentComment.replies ??= []).push(reply);
			continue;
		}
		const comment: Comment = {
			address,
			author,
			text,
			replies: [],
			threadedXml: selfContainedXml(node),
		};
		if (date) comment.date = date;
		const resolved = boolAttr(node, 'done');
		if (resolved !== undefined) comment.resolved = resolved;
		roots.set((att(node, 'id') ?? `${order.length}`).toUpperCase(), comment);
		order.push(comment);
	}
	for (const comment of order)
		comment.replies?.sort((a, b) => {
			const aDate = a.date ? Date.parse(a.date) : NaN;
			const bDate = b.date ? Date.parse(b.date) : NaN;
			return Number.isFinite(aDate) && Number.isFinite(bDate) ? aDate - bDate : 0;
		});
	if (personsXml) {
		const records = new Map<string, readonly [string, string]>();
		for (const person of children(parseXml(personsXml).documentElement, 'person', NS.tc)) {
			const id = att(person, 'id') ?? '';
			records.set(id.toUpperCase(), [id, selfContainedXml(person)]);
		}
		for (const comment of order) {
			const ids = new Set<string>();
			for (const message of [comment, ...(comment.replies ?? [])]) {
				const node = parseXml(message.threadedXml!).documentElement;
				ids.add((att(node, 'personId') ?? '').toUpperCase());
				const mentions = first(node, 'mentions', NS.tc);
				if (mentions)
					for (const mention of children(mentions, 'mention', NS.tc))
						ids.add((att(mention, 'mentionpersonId') ?? '').toUpperCase());
			}
			const kept = [...ids].flatMap((id) => (records.has(id) ? [records.get(id)!] : []));
			if (kept.length) comment.threadedPersons = Object.fromEntries(kept);
		}
	}
	return order;
}

/** Retain the container's unknown attributes and extension list independently of its messages. */
export function threadedCommentEnvelope(xml: string | undefined): string | undefined {
	if (!xml) return undefined;
	const doc = parseXml(xml);
	for (const node of children(doc.documentElement, 'threadedComment', NS.tc)) {
		doc.documentElement.removeChild(node);
	}
	return buildXml(doc);
}

/** Merges threaded comments over the legacy fallbacks Excel writes for them. */
export function mergeComments(legacy: Comment[], threaded: Comment[]): Comment[] {
	if (!threaded.length) return legacy;
	const key = (c: Comment) => `${c.address.row}:${c.address.col}`;
	const threadedAt = new Set(threaded.map(key));
	return [...legacy.filter((c) => !threadedAt.has(key(c))), ...threaded].sort(
		(a, b) => a.address.row - b.address.row || a.address.col - b.address.col,
	);
}
