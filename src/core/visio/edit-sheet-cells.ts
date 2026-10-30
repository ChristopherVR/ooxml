import { executableCellFormula } from './cell-formula';
import { followConnectors } from './edit-instance-geometry';
import { assertNoPageDependents } from './edit-instance-checks';
import { assertOverridable, recalculateInstanceCaches } from './edit-instance-recalculate';
import { instanceScope, singleScope, type InstanceScope } from './edit-instance-scope';
import {
	effectiveNode,
	instanceSheet,
	writeInstanceCell,
	type InstanceCell,
} from './edit-instance-sheet';
import type { MasterTemplate } from './edit-text-instance';
import { visioFormulaCachedValue } from './formula';
import { fail } from './package-common';
import { attribute, children } from './sheet';

const format = (value: number) => String(Object.is(value, -0) ? 0 : value);

/**
 * The sheet an edit of one cell reaches: a stencil instance over its master, or a plain local
 * shape over an empty master. Grouped local shapes are refused.
 */
export async function sheetScope(
	root: Element,
	shapeId: string,
	template: MasterTemplate,
): Promise<InstanceScope> {
	const containers = children(root, 'Shapes');
	const candidates =
		containers.length === 1
			? children(containers[0], 'Shape').filter((shape) => attribute(shape, 'ID') === shapeId)
			: [];
	if (candidates.length !== 1)
		fail('EDIT_TARGET_NOT_FOUND', 'A unique top-level shape is required.');
	const shape = candidates[0]!;
	if (shape.hasAttribute('Master')) {
		const scope = await instanceScope(root, shapeId, template);
		if (!scope) fail('EDIT_TARGET_NOT_FOUND', 'A unique top-level stencil shape is required.');
		return scope!;
	}
	if (shape.hasAttribute('MasterShape'))
		fail('UNSUPPORTED_INSTANCE_EDIT', 'Shapes inside a stencil instance cannot be changed here.');
	return singleScope(
		instanceSheet(shape, shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Shape')),
	);
}

/** The cached number of a cell, or a refusal that names it. */
export function cachedCellNumber(cell: InstanceCell | undefined, what: string): number {
	const node = cell && effectiveNode(cell);
	if (!node || node.hasAttribute('E'))
		fail('UNSUPPORTED_INSTANCE_EDIT', `${what} has no usable value.`);
	try {
		return visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
	} catch {
		return fail('UNSUPPORTED_INSTANCE_EDIT', `${what} is not a number.`);
	}
}

export interface SheetCellWrite {
	cell: InstanceCell;
	value: number;
	/** The formula to store with the value; absent stores a plain value (no F attribute). */
	formula?: string;
}

/**
 * Write cells of a shape the way Visio saves them (local values on the instance), recalculate the
 * caches that follow them, refuse what depends on them elsewhere on the page, and let glued
 * connectors follow. Returns the pages that changed.
 */
export function commitSheetCells(
	roots: ReadonlyMap<string, Element>,
	pageId: string,
	shapeId: string,
	scope: InstanceScope,
	writes: readonly SheetCellWrite[],
	check: () => void,
): readonly string[] {
	const root = roots.get(pageId)!;
	const sheet = scope.root;
	for (const write of writes) assertOverridable(write.cell, write.cell.names[0] ?? 'The cell');
	const overrides = new Map(writes.map((write) => [write.cell, write.value] as const));
	const { writes: caches } = recalculateInstanceCaches(scope, overrides, check);
	const changed = new Map<string, Set<string>>();
	const touch = (owner: typeof sheet, names: readonly string[]) => {
		const id = attribute(owner.instance, 'ID') ?? '';
		const set = changed.get(id) ?? new Set<string>();
		for (const name of names) set.add(name.toLowerCase());
		changed.set(id, set);
	};
	for (const write of writes) touch(scope.owner(write.cell), write.cell.names);
	for (const cache of caches) touch(cache.sheet, cache.cell.names);
	assertNoPageDependents(root, sheet.instance, changed, check);
	for (const write of writes) {
		const created = !write.cell.local;
		const node = writeInstanceCell(scope.owner(write.cell), write.cell, format(write.value));
		// Visio leaves the unit off a cell it creates on the instance.
		if (created) node.removeAttribute('U');
		if (write.formula) node.setAttribute('F', write.formula);
	}
	for (const cache of caches) {
		check();
		const own = cache.cell.local;
		if (own && executableCellFormula(attribute(own, 'F'))) {
			own.setAttribute('V', format(cache.value));
			own.removeAttribute('E');
		} else
			// Unlike a resize, a moved handle leaves the unit tag of a refreshed cache as the master has it.
			writeInstanceCell(cache.sheet, cache.cell, format(cache.value), { formula: 'Inh' });
	}
	return [...new Set([pageId, ...followConnectors(roots, pageId, shapeId, check)])];
}
