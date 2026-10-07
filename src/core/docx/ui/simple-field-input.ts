import { Fragment, Slice, type Mark, type Node } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { TextSelection } from 'prosemirror-state';
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
		if (from !== field.from || to !== field.to || !state.schema.nodes.fieldMarker) return null;
		const run = inlineNodeRun(state.doc.nodeAt(from)!);
		if (!run) return null;
		// A simple field has no separately formatted code runs. Its removed result's
		// direct formatting must not become formatting on the structural markers.
		const format = {
			...(run.commentIds && { commentIds: run.commentIds }),
		};
		const markers = [
			{ ...format, text: '', fieldChar: 'begin' as const },
			{ ...format, text: '', fieldCode: field.mark.attrs.instr as string },
			{ ...format, text: '', fieldChar: 'separate' as const },
			{ ...format, text: '', fieldChar: 'end' as const },
		].flatMap((item) => runToInlineNodes(item, state.schema));
		const tr = state.tr.replaceWith(from, to, markers);
		return tr.setSelection(TextSelection.create(tr.doc, from + 3));
	}
	const tr = state.tr.insertText(text, from, to);
	tr.doc.nodesBetween(from, from + text.length, (node, pos) => {
		if (!node.isText) return;
		for (const mark of resultMarks(node, field))
			tr.addMark(Math.max(from, pos), Math.min(from + text.length, pos + node.nodeSize), mark);
	});
	return tr;
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
