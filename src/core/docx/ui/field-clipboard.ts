import { Fragment, Slice, type Node } from 'prosemirror-model';

/** Result-only copies are literal text; complete complex fields retain their markers and result. */
export function fieldClipboardSlice(slice: Slice): Slice {
	const complete = new Set<number>();
	const stack: number[][] = [];
	slice.content.descendants((node, pos) => {
		if (node.type.name === 'fieldMarker') {
			if (node.attrs.kind === 'begin') stack.push([]);
			else if (node.attrs.kind === 'end') {
				const nodes = stack.pop();
				if (nodes) {
					if (stack.length) stack.at(-1)!.push(...nodes);
					else for (const child of nodes) complete.add(child);
				}
			}
		} else if (node.isText && stack.length) stack.at(-1)!.push(pos);
	});
	const project = (node: Node, pos: number): Node => {
		if (!node.isLeaf) {
			const children: Node[] = [];
			node.forEach((child, offset) => children.push(project(child, pos + 1 + offset)));
			return node.copy(Fragment.fromArray(children));
		}
		const field = node.marks.find((mark) => mark.type.name === 'field');
		if (!field || (!field.attrs.simple && complete.has(pos))) return node;
		const marks = node.marks.flatMap((mark) => {
			if (mark.type === field.type) return [];
			if (mark.type.name !== 'runProperties' || !mark.attrs.props?.fieldInstanceId) return [mark];
			const { fieldInstanceId: _identity, ...props } = mark.attrs.props;
			return Object.keys(props).length ? [mark.type.create({ ...mark.attrs, props })] : [];
		});
		return node.mark(marks);
	};
	const children: Node[] = [];
	slice.content.forEach((node, offset) => children.push(project(node, offset)));
	return new Slice(Fragment.fromArray(children), slice.openStart, slice.openEnd);
}
