import { fail } from './package-common';
import { attribute, children } from './sheet';
import { assertShapeLocks } from './edit-style-admission';

/** Shared source target identity and effective text protection proof. */
export function localTextTarget(
	root: Element,
	document: Element,
	shapeId: string,
	check: () => void,
): Element {
	const shapeChildren = (parent: Element): Element[] => {
		const containers = children(parent, 'Shapes');
		if (containers.length > 1)
			fail('INVALID_SHAPE_ID', 'Duplicate Shapes containers are ambiguous.');
		return children(containers[0], 'Shape');
	};
	const shapes = new Map<string, { node: Element; inherited: boolean; deleted: boolean }>();
	const pending = shapeChildren(root).map((node) => ({ node, inherited: false, deleted: false }));
	while (pending.length) {
		check();
		const item = pending.pop()!,
			id = attribute(item.node, 'ID');
		if (!id || shapes.has(id))
			fail('INVALID_SHAPE_ID', 'Local shape IDs must be present and unique.');
		const inherited =
			item.inherited || item.node.hasAttribute('Master') || item.node.hasAttribute('MasterShape');
		const deleted = item.deleted || ['1', 'true'].includes(attribute(item.node, 'Del') ?? '');
		shapes.set(id, { node: item.node, inherited, deleted });
		for (const node of shapeChildren(item.node)) pending.push({ node, inherited, deleted });
	}
	const target = shapes.get(shapeId);
	if (!target) fail('EDIT_TARGET_NOT_FOUND', 'Local shape does not exist.');
	if (target.inherited || target.deleted)
		fail('UNSUPPORTED_TEXT_EDIT', 'Master-linked or deleted shapes cannot be edited.');
	assertShapeLocks(target.node, document, ['LockTextEdit']);
	return target.node;
}
