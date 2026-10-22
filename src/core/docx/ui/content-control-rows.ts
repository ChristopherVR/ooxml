// Controls around table rows (`w:tbl > w:sdt > w:tr`): the editor keeps a row's control in the
// row's properties, so they are not among `controlSpans`, which the commands operate.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { ContentControl } from '../content-control-model';

export interface RowControlSpan {
	control: ContentControl;
	level: 'row';
	/** The position before the row. */
	from: number;
	/** The position after the row. */
	to: number;
}

/** The control a table row node is wrapped in, if any. */
export const rowControlOf = (node: ProseMirrorNode): ContentControl | undefined =>
	node.type.name === 'tableRow'
		? ((node.attrs.properties as { contentControl?: ContentControl } | null)?.contentControl ??
			undefined)
		: undefined;

/** Every table row wrapped in a control, in document order. */
export function rowControlSpans(doc: ProseMirrorNode): RowControlSpan[] {
	const rows: RowControlSpan[] = [];
	doc.descendants((node, pos) => {
		const control = rowControlOf(node);
		if (control) rows.push({ control, level: 'row', from: pos, to: pos + node.nodeSize });
		return !node.isTextblock;
	});
	return rows;
}
