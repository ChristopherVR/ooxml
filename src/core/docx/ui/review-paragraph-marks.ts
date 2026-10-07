import type { Node } from 'prosemirror-model';
import type { Transaction } from 'prosemirror-state';
import { createDocument, type Paragraph, type Revision } from '../model';
import { resolveParagraphMark } from '../revision-commands';
import { paragraphAttrs, paragraphFromAttrs } from './paragraph-attributes';

export function paragraphMarkRevision(node: Node): Revision | undefined {
	const revision = node.attrs.markRevision as Revision | undefined;
	return node.type.name === 'paragraph' &&
		(revision?.kind === 'insert' || revision?.kind === 'delete')
		? revision
		: undefined;
}

/** Merge model attributes with the existing editor fragments, retaining marks and object identity. */
export function resolveParagraphMarkRange(
	tr: Transaction,
	pos: number,
	mode: 'accept' | 'reject',
): void {
	const node = tr.doc.nodeAt(pos);
	if (!node) return;
	const revision = paragraphMarkRevision(node);
	if (!revision) return;
	const keepBreak = (mode === 'accept') === (revision.kind === 'insert');
	const current = paragraphFromAttrs(node.attrs, String(node.attrs.id), []);
	const nextPos = pos + node.nodeSize;
	const $next = tr.doc.resolve(nextPos);
	const following = $next.nodeAfter;
	if (!keepBreak && typeof tr.doc.attrs.sections === 'string') {
		const sections = JSON.parse(tr.doc.attrs.sections) as { endsAtBlockId?: string }[];
		if (sections.some((section) => section.endsAtBlockId === current.id))
			throw new Error('Cannot resolve a paragraph mark across a section boundary.');
	}
	const blocks = [current];
	if (following?.type.name === 'paragraph')
		blocks.push(paragraphFromAttrs(following.attrs, String(following.attrs.id), []));
	const result = resolveParagraphMark({ ...createDocument(), blocks }, revision.id, keepBreak)
		.blocks[0] as Paragraph;
	if (keepBreak) {
		tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...paragraphAttrs(result) });
		return;
	}
	if (!following || following.type !== node.type)
		throw new Error('Cannot merge different paragraph node types.');
	const merged = following.type.create(
		{ ...following.attrs, ...paragraphAttrs(result) },
		node.content.append(following.content),
		following.marks,
	);
	tr.replaceWith(pos, nextPos + following.nodeSize, merged);
}
