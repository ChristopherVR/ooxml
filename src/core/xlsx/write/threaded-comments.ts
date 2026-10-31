import { NS, buildXml, children, first, parseXml, type XmlDocument } from '../../xml/index';
import type { Comment, CommentMessage } from '../model';
import { CommentIds } from './comment-ids';
import { patchCommentMentions } from './comment-mentions';
import type { PersonRegistry } from './comment-persons';
import { XML_HEADER } from './xml-out';

export interface ThreadedCommentOptions {
	envelope?: string;
	ids?: CommentIds;
}

/** Patch known fields over source messages and retain the part's extension envelope. */
export class ThreadedCommentsWriter {
	private readonly doc: XmlDocument;
	private readonly ids: CommentIds;
	private count = 0;
	constructor(
		private readonly persons: PersonRegistry,
		options: ThreadedCommentOptions = {},
	) {
		this.doc = parseXml(
			options.envelope ?? `${XML_HEADER}<ThreadedComments xmlns="${NS.tc}" xmlns:x="${NS.x}"/>`,
		);
		for (const node of children(this.doc.documentElement, 'threadedComment', NS.tc))
			this.doc.documentElement.removeChild(node);
		this.ids = options.ids ?? new CommentIds();
	}
	add(comment: Comment, ref: string, sheetName: string): string {
		this.persons.retain(comment.threadedPersons);
		const preserve = !comment.regenerateThreadIds;
		const prepare = (message: CommentMessage, seed: string) => {
			const source = parseXml(message.threadedXml ?? `<threadedComment xmlns="${NS.tc}"/>`);
			const priorId = source.documentElement.getAttribute('id') ?? '';
			const id = this.ids.claim(priorId, seed, message, preserve);
			return { message, source, priorId, id };
		};
		const root = prepare(comment, `${sheetName}!${ref}:${comment.text}`);
		const replies = (comment.replies ?? []).map((reply, index) =>
			prepare(reply, `${root.id}:reply:${index}:${reply.text}`),
		);
		const identities = new Map(
			[root, ...replies]
				.filter((item) => item.priorId)
				.map((item) => [item.priorId.toUpperCase(), item.id]),
		);
		this.message(comment, ref, root.source, root.id, preserve, undefined, comment.resolved);
		for (const reply of replies)
			this.message(
				reply.message,
				ref,
				reply.source,
				reply.id,
				preserve,
				identities.get(
					(reply.source.documentElement.getAttribute('parentId') ?? '').toUpperCase(),
				) ?? root.id,
			);
		return root.id;
	}
	private message(
		message: CommentMessage,
		ref: string,
		source: XmlDocument,
		id: string,
		preserve: boolean,
		parent?: string,
		resolved?: boolean,
	): void {
		const node = source.documentElement;
		const personId = this.persons.id(message.author, node.getAttribute('personId') ?? undefined);
		for (const [key, value] of Object.entries({
			id,
			ref,
			personId,
			dT: message.date,
			parentId: parent,
		})) {
			if (value === undefined) node.removeAttribute(key);
			else node.setAttribute(key, value);
		}
		if (!parent) {
			if (resolved === undefined) node.removeAttribute('done');
			else node.setAttribute('done', resolved ? '1' : '0');
		}
		let text = first(node, 'text', NS.tc);
		patchCommentMentions(node, text?.textContent ?? '', message, this.ids, id, preserve);
		if (!text) {
			text = source.createElementNS(NS.tc, node.prefix ? `${node.prefix}:text` : 'text');
			node.insertBefore(text, node.firstChild);
		}
		text.textContent = message.text;
		const root = this.doc.documentElement;
		root.insertBefore(this.doc.importNode(node, true), first(root, 'extLst', NS.tc) ?? null);
		this.count++;
	}
	xml(): string | undefined {
		return this.count ? buildXml(this.doc) : undefined;
	}
}
