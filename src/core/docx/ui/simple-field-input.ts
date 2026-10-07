import { Fragment, Slice, type Mark, type Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';
import type { Transform } from 'prosemirror-transform';
import { fieldResultRanges, type FieldResultRange } from './field-results';
import { inlineNodeRun, runToInlineNodes } from './run-adapter';

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });

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
	const field = targetField(state, from, to);
	if (!field) return null;
	if (!text) {
		const tr = state.tr;
		if (!emptySimpleFieldResult(tr, from, to)) return null;
		return tr.setSelection(TextSelection.create(tr.doc, from + 3));
	}
	const $from = state.doc.resolve(from);
	const marks =
		state.storedMarks ?? (from === to ? $from.marks() : $from.marksAcross(state.doc.resolve(to)));
	const node = state.schema.text(text, marks);
	// Put metadata on the inserted slice before replacing. Adding it afterward emits mark steps,
	// which a text-only Track Changes transaction cannot replay as a tracked replacement.
	const tr = state.tr.replaceRangeWith(from, to, node.mark(resultMarks(node, field)));
	if (!tr.selection.empty && tr.selection.to === from + text.length)
		tr.setSelection(TextSelection.near(tr.selection.$to));
	return tr;
}

/** Removing a complete simple cached result keeps its instruction as an empty complex field. */
export function emptySimpleFieldResult(transform: Transform, from: number, to: number): boolean {
	const field = fieldResultRanges(transform.doc).find(
		(range) => range.mark.attrs.simple && range.from === from && range.to === to,
	);
	const schema = transform.doc.type.schema;
	if (!field || !schema.nodes.fieldMarker) return false;
	const run = inlineNodeRun(transform.doc.nodeAt(from)!);
	if (!run) return false;
	// Result formatting and revisions do not belong to the structural markers.
	const format = { ...(run.commentIds && { commentIds: run.commentIds }) };
	const markers = [
		{ ...format, text: '', fieldChar: 'begin' as const },
		{ ...format, text: '', fieldCode: field.mark.attrs.instr as string },
		{ ...format, text: '', fieldChar: 'separate' as const },
		{ ...format, text: '', fieldChar: 'end' as const },
	].flatMap((item) => runToInlineNodes(item, schema));
	transform.replaceWith(from, to, markers);
	return true;
}

/** Preserve the instruction when Backspace/Delete removes the last cached result text. */
export function deleteSimpleFieldResult(state: EditorState, backward: boolean): Transaction | null {
	const { from, to, empty } = state.selection;
	if (!empty) return replaceSimpleFieldResult(state, from, to, '');
	const field = fieldResultRanges(state.doc).find(
		(range) => range.mark.attrs.simple && (backward ? range.to === from : range.from === from),
	);
	if (!field) return null;
	// Positions count UTF-16 units, whereas a deletion can remove one grapheme
	// containing several units, even when direct formatting splits its runs.
	const segments = graphemes.segment(field.text)[Symbol.iterator]();
	if (segments.next().done || !segments.next().done) return null;
	return replaceSimpleFieldResult(state, field.from, field.to, '');
}
