import { Fragment, Slice, type Node } from 'prosemirror-model';
import { inlineRunProperties, updatedInlineRunAttributes } from './inline-run-properties';

/** Incomplete field fragments become literals; complete complex fields retain their markers and result. */
export function fieldClipboardSlice(slice: Slice): Slice {
	const complete = new Set<number>();
	const stack: { positions: number[]; result: boolean; valid: boolean }[] = [];
	slice.content.descendants((node, pos) => {
		if (node.type.name === 'fieldMarker') {
			if (node.attrs.kind === 'begin') stack.push({ positions: [pos], result: false, valid: true });
			else if (node.attrs.kind === 'end') {
				const open = stack.pop();
				if (open?.valid) {
					complete.add(pos);
					for (const child of open.positions) complete.add(child);
				}
			} else {
				const open = stack.at(-1);
				if (open) {
					open.positions.push(pos);
					if (node.attrs.kind === 'code') open.valid &&= !open.result;
					else if (node.attrs.kind === 'separate') {
						open.valid &&= !open.result;
						open.result = true;
					} else open.valid = false;
				}
			}
		} else if (node.isInline && node.isLeaf && stack.length) stack.at(-1)!.positions.push(pos);
	});
	const project = (node: Node, pos: number): Node | null => {
		if (node.type.name === 'fieldMarker' && !complete.has(pos)) return null;
		if (!node.isLeaf) {
			const children: Node[] = [];
			node.forEach((child, offset) => {
				const projected = project(child, pos + 1 + offset);
				if (projected) children.push(projected);
			});
			return node.copy(Fragment.fromArray(children));
		}
		const properties =
			!node.isText && node.type.spec.attrs?.format ? inlineRunProperties(node) : undefined;
		const field = node.marks.find((mark) => mark.type.name === 'field');
		const cached = properties?.field;
		if (!field && !cached) return node;
		if (!(field?.attrs.simple ?? cached?.simple) && complete.has(pos)) return node;
		const marks = node.marks.flatMap((mark) => {
			if (mark.type === field?.type) return [];
			if (mark.type.name !== 'runProperties') return [mark];
			const { fieldInstanceId: _identity, fieldFlags: _flags, ...props } = mark.attrs.props ?? {};
			return Object.keys(props).length ? [mark.type.create({ ...mark.attrs, props })] : [];
		});
		if (!properties || !cached) return node.mark(marks);
		const { field: _field, fieldInstanceId: _id, fieldFlags: _flags, ...literal } = properties;
		return node.type.create(
			updatedInlineRunAttributes(
				node,
				Object.keys(literal).length ? JSON.stringify(literal) : null,
			),
			node.content,
			marks,
		);
	};
	const children: Node[] = [];
	slice.content.forEach((node, offset) => {
		const projected = project(node, offset);
		if (projected) children.push(projected);
	});
	return new Slice(Fragment.fromArray(children), slice.openStart, slice.openEnd);
}
