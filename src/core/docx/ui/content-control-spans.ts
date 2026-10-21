// Where content controls sit in the editor document: inline controls are runs of text sharing a
// `contentControl` mark, block controls are consecutive paragraphs sharing a node attribute, and
// cell controls are a table cell's node attribute. A block control around a table is the table's
// node attribute; no command operates it, so it is listed apart (`tableControlSpans`).
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

/**
 * The control `data` (a `data-sdt` attribute, or an inline control's mark) holds, at `level`. HTML
 * pasted from another page and remote peers can put anything there: only an object with an id and
 * a kind is a control.
 */
export function controlFromData(
	data: unknown,
	level: ContentControl['level'],
): ContentControl | undefined {
	if (!data) return undefined;
	let value: unknown;
	try {
		value = JSON.parse(String(data));
	} catch {
		return undefined;
	}
	return controlFromValue(value, level);
}

/**
 * The control `value` is at `level`: itself when it is an object with an id and a kind at that
 * level, a copy at that level when it is at another, nothing when it is not a control.
 */
export function controlFromValue(
	value: unknown,
	level: ContentControl['level'],
): ContentControl | undefined {
	if (!value || typeof value !== 'object' || Array.isArray(value)) return undefined;
	const { id, kind } = value as { id?: unknown; kind?: unknown };
	if (typeof id !== 'string' || typeof kind !== 'string') return undefined;
	const control = value as ContentControl;
	return control.level === level ? control : { ...control, level };
}

export const controlOfMark = (mark: Mark | undefined): ContentControl | undefined =>
	controlFromData(mark?.attrs.data, 'inline');

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

/** The tables held by a block control, in document order. */
export function tableControlSpans(doc: ProseMirrorNode): ControlSpan[] {
	const spans: ControlSpan[] = [];
	doc.descendants((node, pos) => {
		const control = node.type.name === 'table' ? node.attrs.contentControl : undefined;
		if (control)
			spans.push({
				control: control as ContentControl,
				level: 'block',
				from: pos,
				to: pos + node.nodeSize,
			});
		return !node.isTextblock;
	});
	return spans;
}

/**
 * The innermost control around `pos`. An inline control wins over the block or cell around it,
 * and with `edges` also when `pos` is at its start or end.
 */
export function controlAt(
	spans: readonly ControlSpan[],
	pos: number,
	edges = false,
): ControlSpan | undefined {
	const inside = spans.filter((span) => pos > span.from && pos < span.to);
	const inline =
		inside.find((span) => span.level === 'inline') ??
		(edges
			? spans.find((span) => span.level === 'inline' && (pos === span.from || pos === span.to))
			: undefined);
	return (
		inline ??
		inside.reduce<ControlSpan | undefined>(
			(inner, span) => (!inner || span.to - span.from < inner.to - inner.from ? span : inner),
			undefined,
		)
	);
}

export const controlAtSelection = (state: EditorState, edges = true): ControlSpan | undefined =>
	controlAt(controlSpans(state.doc), state.selection.from, edges);
