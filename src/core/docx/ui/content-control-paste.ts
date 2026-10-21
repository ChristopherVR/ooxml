// A pasted content control gets a new id when the document already has a control with its id.
// Ids are unique in a document (ECMA-376 17.5.2.18: Word gives a repeated one a new id when it
// opens the file), and a copy pasted beside its source must be operated and saved as itself.
// Text pasted inside a control joins it, as typed text does, so that control keeps its id. Block
// controls keep their ids: the writer joins a new block to the control of the block before it by
// id, as it does for a paragraph typed at the end of one.
import { Fragment, Slice, type Mark, type Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { freshControlId, type ContentControl } from '../content-control-model';
import { rowControlOf, rowControlSpans } from './content-control-rows';
import { controlOfMark, controlSpans, tableControlSpans } from './content-control-spans';

/**
 * `slice` as it is pasted in `view`: its inline, cell and row controls under ids the document
 * lacks. Plain text (`plain`) takes the marks of where it lands, and a drag that moves content
 * takes its controls along: neither holds a copy of a control.
 */
export function pastedControlSlice(
	slice: Slice,
	view: EditorView | undefined,
	plain = false,
): Slice {
	if (!view || plain || view.dragging?.move) return slice;
	const { doc, selection, storedMarks } = view.state;
	const spans = [...controlSpans(doc), ...tableControlSpans(doc), ...rowControlSpans(doc)];
	const inDocument = new Set(spans.map((span) => span.control.id));
	// The inline controls pasted text joins: those whose mark text typed there would take. A
	// drag's selection is where it came from, not where it lands, so a dropped copy joins none.
	const landing = view.dragging
		? []
		: selection.empty
			? (storedMarks ?? selection.$from.marks())
			: (selection.$from.marksAcross(selection.$to) ?? []);
	const joined = new Set(
		landing.flatMap((mark) => {
			const id = mark.type.name === 'contentControl' ? controlOfMark(mark)?.id : undefined;
			return id ? [id] : [];
		}),
	);
	// A new id is new to the document and to the pasted content both.
	const taken = new Set(inDocument);
	slice.content.descendants((node) => {
		for (const mark of node.marks) {
			const id = mark.type.name === 'contentControl' ? controlOfMark(mark)?.id : undefined;
			if (id) taken.add(id);
		}
		const id = (node.attrs.contentControl as ContentControl | null | undefined)?.id;
		const row = rowControlOf(node)?.id;
		for (const found of [id, row]) if (found) taken.add(found);
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
		const copied = mark.type.name === 'contentControl' ? controlOfMark(mark) : undefined;
		const control = copied && !joined.has(copied.id) ? withNewId(copied) : undefined;
		return control ? mark.type.create({ ...mark.attrs, data: JSON.stringify(control) }) : mark;
	};
	const nodeOf = (node: ProseMirrorNode): ProseMirrorNode => {
		if (node.isInline) {
			const marks = node.marks.map(markOf);
			return marks.some((mark, index) => mark !== node.marks[index]) ? node.mark(marks) : node;
		}
		const content = fragmentOf(node.content);
		const cell =
			node.type.name === 'tableCell'
				? withNewId(node.attrs.contentControl as ContentControl | undefined)
				: undefined;
		const row = withNewId(rowControlOf(node));
		if (!cell && !row && content === node.content) return node;
		const attrs = cell
			? { ...node.attrs, contentControl: cell }
			: row
				? { ...node.attrs, properties: { ...node.attrs.properties, contentControl: row } }
				: node.attrs;
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
