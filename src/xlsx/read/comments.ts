import { NS, children, parseXml } from '../../xml/index.js';
import { parseAddress } from '../address.js';
import type { Comment } from '../model.js';
import { att, xChildren, xFirst, xText } from './xml-util.js';

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
): Comment[] {
	if (!xml) return [];
	const root = parseXml(xml, { label: 'XLSX threaded comments' }).documentElement;
	const roots = new Map<string, Comment>();
	const order: Comment[] = [];
	for (const node of children(root, 'threadedComment', NS.tc)) {
		const address = parseAddress(att(node, 'ref') ?? '');
		if (!address) continue;
		const author = persons.get(att(node, 'personId') ?? '') ?? '';
		const text = xText(children(node, 'text', NS.tc)[0]);
		const date = att(node, 'dT');
		const parent = att(node, 'parentId');
		const parentComment = parent ? roots.get(parent) : undefined;
		if (parentComment) {
			const reply: { author: string; text: string; date?: string } = { author, text };
			if (date) reply.date = date;
			(parentComment.replies ??= []).push(reply);
			continue;
		}
		const comment: Comment = { address, author, text, replies: [] };
		roots.set(att(node, 'id') ?? `${order.length}`, comment);
		order.push(comment);
	}
	return order;
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
