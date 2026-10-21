// Where content controls sit in the editor document: inline controls are runs of text sharing a
// `contentControl` mark, block controls are consecutive paragraphs sharing a node attribute, and
// cell controls are a table cell's node attribute.
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { ContentControl } from '../content-control-model';

/** A control and the document range it covers. */
export interface ControlSpan {
	control: ContentControl;
	level: 'inline' | 'block' | 'cell';
	/** Inline: the first text position. Block or cell: the position before its first node. */
	from: number;
	/** Inline: after the last character. Block or cell: after its last node. */
	to: number;
}

export function controlOfMark(mark: Mark | undefined): ContentControl | undefined {
	if (!mark?.attrs.data) return undefined;
	try {
		return JSON.parse(String(mark.attrs.data)) as ContentControl;
	} catch {
		return undefined;
	}
}

const markOf = (node: ProseMirrorNode): Mark | undefined =>
	node.marks.find((mark) => mark.type.name === 'contentControl');

/** Every inline, block and cell control in document order. */
export function controlSpans(doc: ProseMirrorNode): ControlSpan[] {
	const spans: ControlSpan[] = [];
	// The last block span: the paragraphs of one block control stay one span when an inline
	// control in one of them comes between them in this list.
	let block: ControlSpan | undefined;
	doc.descendants((node, pos) => {
		if (node.type.name === 'tableCell' && node.attrs.contentControl)
			spans.push({
				control: node.attrs.contentControl as ContentControl,
				level: 'cell',
				from: pos,
				to: pos + node.nodeSize,
			});
		if (node.type.name === 'paragraph' && node.attrs.contentControl) {
			const control = node.attrs.contentControl as ContentControl;
			if (block?.control.id === control.id && block.to === pos) block.to = pos + node.nodeSize;
			else {
				block = { control, level: 'block', from: pos, to: pos + node.nodeSize };
				spans.push(block);
			}
		}
		if (!node.isInline) return true;
		const control = controlOfMark(markOf(node));
		if (!control) return false;
		const previous = spans.at(-1);
		if (previous?.level === 'inline' && previous.control.id === control.id && previous.to === pos) {
			previous.to = pos + node.nodeSize;
			previous.control = control;
		} else spans.push({ control, level: 'inline', from: pos, to: pos + node.nodeSize });
		return false;
	});
	return spans;
}

/** The innermost control around `pos` (an inline control wins over the block or cell around it). */
export function controlAt(
	spans: readonly ControlSpan[],
	pos: number,
	edges = false,
): ControlSpan | undefined {
	const inside = spans.filter((span) => pos > span.from && pos < span.to);
	const found =
		inside.find((span) => span.level === 'inline') ??
		inside.reduce<ControlSpan | undefined>(
			(inner, span) => (!inner || span.to - span.from < inner.to - inner.from ? span : inner),
			undefined,
		);
	if (found || !edges) return found;
	return spans.find((span) => span.level === 'inline' && (pos === span.from || pos === span.to));
}

export const controlAtSelection = (state: EditorState, edges = true): ControlSpan | undefined =>
	controlAt(controlSpans(state.doc), state.selection.from, edges);
