// A pasted content control gets a new id when the document already has a control with its id.
// Ids are unique in a document (ECMA-376 17.5.2.18: Word gives a repeated one a new id when it
// opens the file), and a copy pasted beside its source must be operated and saved as itself.
// Block controls keep their ids: the writer joins a new block to the control of the block before
// it by id, as it does for a paragraph typed at the end of one.
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { ContentControl } from '../content-control-model';
import { controlOfMark, controlSpans } from './content-control-spans';

/** A control id that is not in `used`: Word uses signed 32-bit integers. */
export function freshControlId(used: ReadonlySet<string>): string {
	let id: string;
	do id = String(Math.floor(Math.random() * 2 ** 31) + 1);
	while (used.has(id));
	return id;
}

/** `slice` as it is pasted in `view`: its inline and cell controls under ids the document lacks. */
export function pastedControlSlice(slice: Slice, view: EditorView | undefined): Slice {
	// A drag that moves content takes its controls along, ids and all.
	if (!view || view.dragging?.move) return slice;
	const inDocument = new Set(controlSpans(view.state.doc).map((span) => span.control.id));
	// A new id is new to the document and to the pasted content both.
	const taken = new Set(inDocument);
	slice.content.descendants((node) => {
		for (const mark of node.marks) {
			const id = mark.type.name === 'contentControl' ? controlOfMark(mark)?.id : undefined;
			if (id) taken.add(id);
		}
		const id = (node.attrs.contentControl as ContentControl | null | undefined)?.id;
		if (id) taken.add(id);
	});
	const renamed = new Map<string, string>();
	const withNewId = (control: ContentControl | undefined): ContentControl | undefined => {
		if (!control || !inDocument.has(control.id)) return undefined;
		let id = renamed.get(control.id);
		if (!id) {
			id = freshControlId(taken);
			taken.add(id);
			renamed.set(control.id, id);
		}
		return { ...control, id };
	};
	const markOf = (mark: Mark): Mark => {
		const control =
			mark.type.name === 'contentControl' ? withNewId(controlOfMark(mark)) : undefined;
		return control ? mark.type.create({ ...mark.attrs, data: JSON.stringify(control) }) : mark;
	};
	const nodeOf = (node: ProseMirrorNode): ProseMirrorNode => {
		if (node.isInline) {
			const marks = node.marks.map(markOf);
			return marks.some((mark, index) => mark !== node.marks[index]) ? node.mark(marks) : node;
		}
		const content = fragmentOf(node.content);
		const control =
			node.type.name === 'tableCell'
				? withNewId(node.attrs.contentControl as ContentControl | undefined)
				: undefined;
		if (!control && content === node.content) return node;
		const attrs = control ? { ...node.attrs, contentControl: control } : node.attrs;
		return node.type.create(attrs, content, node.marks);
	};
	const fragmentOf = (fragment: Fragment): Fragment => {
		const nodes: ProseMirrorNode[] = [];
		let changed = false;
		fragment.forEach((node) => {
			const next = nodeOf(node);
			changed ||= next !== node;
			nodes.push(next);
		});
		return changed ? Fragment.fromArray(nodes) : fragment;
	};
	const content = fragmentOf(slice.content);
	return content === slice.content ? slice : new Slice(content, slice.openStart, slice.openEnd);
}
