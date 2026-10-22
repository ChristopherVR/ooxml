// One content control in the editor document can be several spans: the paragraphs of a block
// control with a table between them, the rows or cells of one `w:sdt` around several, the text of an
// inline control around a nested one. Its occurrence is all of them, from its first to its last.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { RowControlSpan } from './content-control-rows';
import type { ControlSpan } from './content-control-spans';

export type Span = ControlSpan | RowControlSpan;

/** One control and its range: its spans, more than one when it holds what splits them. */
export interface Occurrence<S extends Span = Span> {
	key: string;
	spans: S[];
	from: number;
	to: number;
}

/**
 * Whether `span` goes on with the control of its level and id that ends at `end`: right after
 * it, or past only what splits one control, other inline controls in its text (a nested control)
 * and, among its paragraphs, blocks other than paragraphs (a table) and paragraphs of another
 * control (a nested block control).
 */
function continues(
	doc: ProseMirrorNode,
	inline: readonly Span[],
	end: number,
	span: Span,
): boolean {
	if (span.from === end) return true;
	if (span.level === 'inline') {
		let at = end;
		for (const other of inline) if (other.from === at && other.to <= span.from) at = other.to;
		return at === span.from;
	}
	if (span.level !== 'block') return false;
	const $end = doc.resolve(end);
	if (doc.resolve(span.from).parent !== $end.parent) return false;
	let split = true;
	$end.parent.forEach((node, offset) => {
		const pos = $end.start() + offset;
		if (pos < end || pos >= span.from || node.type.name !== 'paragraph') return;
		if (!node.attrs.contentControl) split = false;
	});
	return split;
}

/**
 * The occurrences of the controls `spans` hold, in document order: the spans of one level and id
 * that go on with each other are one. `inline` is every inline span of `doc`, which can split one.
 */
export function controlOccurrences<S extends Span>(
	doc: ProseMirrorNode,
	spans: readonly S[],
	inline: readonly Span[],
): Occurrence<S>[] {
	const found: Occurrence<S>[] = [];
	const last = new Map<string, Occurrence<S>>();
	for (const span of [...spans].sort((a, b) => a.from - b.from)) {
		const key = `${span.level}:${span.control.id}`;
		const previous = last.get(key);
		if (previous && continues(doc, inline, previous.to, span)) {
			previous.spans.push(span);
			previous.to = span.to;
			continue;
		}
		const occurrence = { key, spans: [span], from: span.from, to: span.to };
		found.push(occurrence);
		last.set(key, occurrence);
	}
	return found;
}
