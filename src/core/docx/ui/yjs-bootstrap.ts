import type { Node } from 'prosemirror-model';
import * as Y from 'yjs';
import { commentIdsFromMarks, commentIdsFromNode } from './comment-anchors';
import { inlineRunProperties, updatedInlineRunAttributes } from './inline-run-properties';

/** Normalize loaded legacy grouped marks before creating a new shared room. */
export function independentCommentAnchors(node: Node): Node {
	const comment = node.type.schema.marks.comment;
	const marks = comment
		? [
				...node.marks.filter((mark) => mark.type !== comment),
				...commentIdsFromMarks(node.marks).map((id) => comment.create({ ids: [id] })),
			]
		: node.marks;
	if (node.isLeaf) {
		if (
			!node.isText &&
			node.type.spec.attrs?.format &&
			node.marks.some(
				(mark) => mark.type.name === 'comment' || mark.type.name === 'inlineCommentAnchors',
			)
		) {
			const properties = inlineRunProperties(node);
			const ids = commentIdsFromNode(node);
			if (ids.length) properties.commentIds = ids;
			else delete properties.commentIds;
			return node.type.create(
				updatedInlineRunAttributes(node, JSON.stringify(properties)),
				undefined,
				marks.filter(
					(mark) => mark.type.name !== 'inlineCommentAnchors' && mark.type.name !== 'comment',
				),
			);
		}
		return node.mark(marks);
	}
	const children: Node[] = [];
	node.forEach((child) => children.push(independentCommentAnchors(child)));
	return node.type.create(node.attrs, children, marks);
}

/** A shared empty text container avoids concurrent first-insertion text-node merging
 * in y-prosemirror (issue 160), which otherwise loses the original undo ownership. */
export function seedEmptyParagraphText(fragment: Y.XmlFragment): void {
	for (const child of fragment.toArray()) {
		if (!(child instanceof Y.XmlElement)) continue;
		if (child.nodeName === 'paragraph' && child.length === 0) child.insert(0, [new Y.XmlText()]);
		else seedEmptyParagraphText(child);
	}
}
