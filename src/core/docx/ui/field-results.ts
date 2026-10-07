import type { Mark, Node } from 'prosemirror-model';

export interface FieldResultRange {
	from: number;
	to: number;
	mark: Mark;
	text: string;
	marks: readonly Mark[];
}

/** Adjacent text carrying the same field mark, confined to its paragraph. */
export function fieldResultRanges(doc: Node): FieldResultRange[] {
	const runs: FieldResultRange[] = [];
	const field = doc.type.schema.marks.field;
	if (!field) return runs;
	doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		let open: FieldResultRange | undefined;
		node.forEach((child, offset) => {
			const mark = child.isText ? child.marks.find((item) => item.type === field) : undefined;
			const start = pos + 1 + offset;
			if (mark && open?.mark.eq(mark)) {
				open.to = start + child.nodeSize;
				open.text += child.text ?? '';
			} else if (mark) {
				open = {
					from: start,
					to: start + child.nodeSize,
					mark,
					text: child.text ?? '',
					marks: child.marks,
				};
				runs.push(open);
			} else open = undefined;
		});
		return false;
	});
	return runs;
}
