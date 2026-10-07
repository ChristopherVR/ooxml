import { Fragment, Slice, type Node } from 'prosemirror-model';
import type { Transform } from 'prosemirror-transform';
import { fieldResultRanges } from './field-results';
import { fieldClipboardSlice } from './field-clipboard';
import { inlineTextRevision } from './review-inline-revisions';
import { expectDefined } from './defined';

/** A structural empty-result replacement keeps tracked text inside its one field. */
export function trackSimpleFieldDeletion(
	transform: Transform,
	from: number,
	to: number,
	slice: Slice,
	author: string,
	date: string,
	nextRevisionId: () => string,
): { deletionIds: string[]; deletedText: string; insertedText: string } | null {
	if (slice.openStart || slice.openEnd || slice.content.childCount !== 4) return null;
	const kinds = ['begin', 'code', 'separate', 'end'];
	for (let index = 0; index < kinds.length; index++) {
		const node = slice.content.child(index);
		if (node.type.name !== 'fieldMarker' || node.attrs.kind !== kinds[index]) return null;
	}
	const field = fieldResultRanges(transform.doc).find(
		(range) => range.from === from && range.to === to && range.mark.attrs.simple,
	);
	if (!field || slice.content.child(1).attrs.code !== field.mark.attrs.instr) return null;
	const result = transform.doc.slice(from, to).content;
	const kept: Node[] = [];
	result.forEach((node) => {
		const revision = inlineTextRevision(node);
		if (!revision || !['insert', 'moveTo'].includes(revision.kind) || revision.author !== author)
			kept.push(node);
	});
	const deletionIds: string[] = [];
	const deletedText = kept.map((node) => node.text ?? '').join('');
	let content = Fragment.empty;
	if (kept.length) {
		const id = nextRevisionId();
		deletionIds.push(id);
		const deletion = expectDefined(
			transform.doc.type.schema.marks.deletion,
			'deletion mark',
		).create({ author, date, id });
		const mark = field.mark.type.create({ ...field.mark.attrs, simple: false });
		content = fieldClipboardSlice(new Slice(Fragment.fromArray(kept), 0, 0)).content;
		const nodes: Node[] = [];
		content.forEach((node) => nodes.push(node.mark(deletion.addToSet(mark.addToSet(node.marks)))));
		content = Fragment.fromArray(nodes);
	}
	transform.replaceWith(
		from,
		to,
		slice.content
			.cut(0, 3)
			.append(content)
			.append(Fragment.from(slice.content.child(3))),
	);
	return { deletionIds, deletedText, insertedText: '' };
}
