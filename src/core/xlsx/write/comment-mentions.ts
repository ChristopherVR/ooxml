import { diffText } from '../../text/text-diff';
import { NS, children, first, type XmlElement } from '../../xml/index';
import type { CommentMessage } from '../model';
import type { CommentIds } from './comment-ids';

/** Keep mentions whose text survives; move their UTF-16 offsets with edits around them. */
export function patchCommentMentions(
	node: XmlElement,
	before: string,
	message: CommentMessage,
	ids: CommentIds,
	commentId: string,
	preserveIds: boolean,
): void {
	const container = first(node, 'mentions', NS.tc);
	if (!container) return;
	const changes = diffText(before, message.text);
	children(container, 'mention', NS.tc).forEach((mention, index) => {
		const start = Number(mention.getAttribute('startIndex'));
		const length = Number(mention.getAttribute('length'));
		const end = start + length;
		let shift = 0;
		let intact =
			Number.isInteger(start) &&
			start >= 0 &&
			Number.isInteger(length) &&
			length > 0 &&
			end <= before.length;
		for (const change of changes) {
			if (change.fromEnd <= start)
				shift += change.toEnd - change.toStart - (change.fromEnd - change.fromStart);
			else if (change.fromStart < end) intact = false;
		}
		if (!intact) {
			container.removeChild(mention);
			return;
		}
		mention.setAttribute('startIndex', String(start + shift));
		mention.setAttribute(
			'mentionId',
			ids.claim(
				mention.getAttribute('mentionId') ?? '',
				`${commentId}:mention:${index}`,
				ids.mentionOwner(message, index),
				preserveIds,
			),
		);
	});
	if (!children(container, 'mention', NS.tc).length) node.removeChild(container);
}
