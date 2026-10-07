import { Fragment, Slice, type Mark, type Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { fieldResultRanges, type FieldResultRange } from './field-results';

function targetField(state: EditorState, from: number, to: number): FieldResultRange | undefined {
	return fieldResultRanges(state.doc).find(
		(range) =>
			range.mark.attrs.simple &&
			from >= range.from &&
			to <= range.to &&
			(from !== to || (from > range.from && from < range.to)),
	);
}

function resultMarks(node: Node, field: FieldResultRange): readonly Mark[] {
	let marks = field.mark.addToSet(node.marks);
	const id = field.marks.find((mark) => mark.type.name === 'runProperties')?.attrs.props
		?.fieldInstanceId;
	const properties = node.type.schema.marks.runProperties;
	if (typeof id === 'string' && properties) {
		const old = properties.isInSet(marks);
		marks = properties
			.create({
				...(old?.attrs ?? {}),
				props: { ...(old?.attrs.props ?? {}), fieldInstanceId: id },
			})
			.addToSet(marks);
	}
	return marks;
}

/** Inline pasted text adopts the existing simple result; boundary cursors and structural slices do not. */
export function simpleFieldPasteSlice(slice: Slice, state: EditorState | undefined): Slice {
	if (!state) return slice;
	const field = targetField(state, state.selection.from, state.selection.to);
	if (!field) return slice;
	const first = slice.content.firstChild;
	if (first?.type.name === 'paragraph' && (slice.openStart === 0 || slice.openEnd === 0))
		return slice;
	const inline =
		first?.type.name === 'paragraph' && slice.content.childCount === 1
			? first.content
			: slice.content;
	if (
		!inline.childCount ||
		!Array.from({ length: inline.childCount }, (_, index) => inline.child(index)).every(
			(node) => node.isText,
		)
	)
		return slice;
	const project = (node: Node): Node => {
		if (node.isText) return node.mark(resultMarks(node, field));
		const children: Node[] = [];
		node.forEach((child) => children.push(project(child)));
		return node.copy(Fragment.fromArray(children));
	};
	const children: Node[] = [];
	slice.content.forEach((node) => children.push(project(node)));
	return new Slice(Fragment.fromArray(children), slice.openStart, slice.openEnd);
}

/** Typed replacements retain the instruction and boundary identity of the cached result. */
export function replaceSimpleFieldResult(
	state: EditorState,
	from: number,
	to: number,
	text: string,
): Transaction | null {
	const field = text ? targetField(state, from, to) : undefined;
	if (!field) return null;
	const tr = state.tr.insertText(text, from, to);
	tr.doc.nodesBetween(from, from + text.length, (node, pos) => {
		if (!node.isText) return;
		for (const mark of resultMarks(node, field))
			tr.addMark(Math.max(from, pos), Math.min(from + text.length, pos + node.nodeSize), mark);
	});
	return tr;
}
