import type { VisioEdit } from './edit-commands';
import { assertInstanceLayersUnlocked } from './edit-instance-format';
import type { VisioPackage } from './package';
import { markPlannedMemberCell, markStencilMember } from './edit-group-rotation';
import { analyzeVisioFormula } from './formula';
import { executableCellFormula } from './cell-formula';
import { assertNoPageDependents } from './edit-instance-checks';
import { writeInstanceResult } from './edit-instance-geometry';
import { assertOverridable, recalculateInstanceCaches } from './edit-instance-recalculate';
import { isGroupInstance, singleScope, type InstanceScope } from './edit-instance-scope';
import {
	effectiveFormula,
	effectiveNode,
	instanceSheet,
	type InstanceCell,
	type InstanceSheet,
} from './edit-instance-sheet';
import type { MasterTemplate } from './edit-text-instance';
import { visioFormulaCachedValue } from './formula';
import { fail, VisioPackageError } from './package-common';
import { attribute, children } from './sheet';

/** A stencil (master) instance that is a member of a group being resized. */
export interface GroupMemberInstance {
	sheet: InstanceSheet;
	scope: InstanceScope;
}
const prepared = new WeakMap<Element, GroupMemberInstance>();
export const groupMemberInstance = (shape: Element): GroupMemberInstance | undefined =>
	prepared.get(shape);

/**
 * Resolve the masters of the stencil members inside the groups that `commands` resize, before the
 * synchronous edit runs (the master part is read asynchronously). Single-shape masters only: a
 * member that is itself a group instance keeps its own scaling, which is not reproduced here.
 */
export async function prepareGroupResizeInstances(
	roots: ReadonlyMap<string, Element>,
	commands: readonly VisioEdit[],
	template: MasterTemplate,
	check: () => void,
): Promise<void> {
	for (const command of commands) {
		if (command.type !== 'resize-shape') continue;
		const root = roots.get(command.pageId);
		const top = root
			? children(children(root, 'Shapes')[0], 'Shape').find(
					(shape) => attribute(shape, 'ID') === command.shapeId,
				)
			: undefined;
		if (!top || attribute(top, 'Type') !== 'Group') continue;
		const pending = [top];
		while (pending.length) {
			check();
			const parent = pending.pop()!;
			for (const container of children(parent, 'Shapes'))
				for (const shape of children(container, 'Shape')) {
					if (!shape.hasAttribute('Master')) {
						pending.push(shape);
						continue;
					}
					if (prepared.has(shape)) continue;
					if (isGroupInstance(shape))
						fail('UNSUPPORTED_GROUP_EDIT', 'A member is an instance of a group master.');
					let base: Element;
					try {
						base = await template(attribute(shape, 'Master')!);
					} catch (error) {
						if (!(error instanceof VisioPackageError) || error.code.startsWith('LIMIT_'))
							throw error;
						return fail(
							'UNSUPPORTED_GROUP_EDIT',
							'The master of a stencil member cannot be resolved.',
						);
					}
					const sheet = instanceSheet(shape, base);
					const parentId = attribute(parent, 'ID');
					for (const name of ['pinx', 'piny', 'width', 'height', 'locpinx', 'locpiny']) {
						const local = sheet.byName.get(name)?.local;
						const source = local && executableCellFormula(attribute(local, 'F'));
						// Only formulas of the group's own size (or the member's) are rewritten here.
						if (
							source &&
							analyzeVisioFormula(source).references.some(
								(ref) =>
									!['width', 'height'].includes(ref.cell.toLowerCase()) ||
									(ref.shapeId !== undefined && ref.shapeId !== parentId),
							)
						)
							fail(
								'UNSUPPORTED_GROUP_EDIT',
								'A stencil member keeps its own formula for its pose.',
							);
						if (local) markPlannedMemberCell(local);
					}
					prepared.set(shape, { sheet, scope: singleScope(sheet) });
					markStencilMember(shape);
				}
		}
	}
}

/**
 * A group resize is refused while any member sits on a locked layer, as the same resize of that
 * member on its own would be. Layer membership does not otherwise touch geometry.
 */
export async function assertGroupMemberLayersUnlocked(
	pkg: VisioPackage,
	roots: ReadonlyMap<string, Element>,
	commands: readonly VisioEdit[],
	check: () => void,
): Promise<void> {
	for (const command of commands) {
		if (command.type !== 'resize-shape') continue;
		const root = roots.get(command.pageId);
		const top = root
			? children(children(root, 'Shapes')[0], 'Shape').find(
					(shape) => attribute(shape, 'ID') === command.shapeId,
				)
			: undefined;
		if (!top || attribute(top, 'Type') !== 'Group') continue;
		const pending = [top];
		while (pending.length) {
			check();
			for (const container of children(pending.pop()!, 'Shapes'))
				for (const shape of children(container, 'Shape')) {
					pending.push(shape);
					const layers = children(shape, 'Cell').find(
						(cell) => attribute(cell, 'N') === 'LayerMember',
					);
					await assertInstanceLayersUnlocked(pkg, command.pageId, attribute(layers, 'V'));
				}
		}
	}
}

const close = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-10 * Math.max(1, Math.abs(a), Math.abs(b));

/** A cell in effect over the master, as a number. */
export function instanceNumber(inst: GroupMemberInstance, name: string, fallback?: number): number {
	const item = inst.sheet.byName.get(name.toLowerCase());
	const node = item && effectiveNode(item);
	if (!node) {
		if (fallback !== undefined) return fallback;
		return fail('UNSUPPORTED_GROUP_EDIT', `A stencil member has no ${name}.`);
	}
	if (node.hasAttribute('E'))
		fail('UNSUPPORTED_GROUP_EDIT', `${name} of a stencil member is an error.`);
	try {
		return visioFormulaCachedValue(attribute(node, 'V') ?? '', attribute(node, 'U')).value;
	} catch {
		return fail('UNSUPPORTED_GROUP_EDIT', `${name} of a stencil member is not a number.`);
	}
}

/** The six transform cells a group resize scales, with the proportional formula Visio writes. */
export interface InstanceMemberPose {
	/** Cell name to the new value; `formula` is the proportional `Sheet.N!Width*k` text. */
	cells: { name: string; value: number; formula: string }[];
	width: number;
	height: number;
}

/**
 * Plan one stencil member's part of a group resize: PinX, PinY, Width and Height become local
 * proportional formulas of the group (as Visio writes for any member), LocPin follows the master's
 * own formula when it has one, and the inherited caches (geometry rows, connection points, text
 * block) are refreshed. Nothing is written until the returned function runs, after every proof.
 */
export function planInstanceMember(
	root: Element,
	shape: Element,
	inst: GroupMemberInstance,
	pose: InstanceMemberPose,
	check: () => void,
): () => void {
	const cell = (name: string) => inst.sheet.byName.get(name.toLowerCase());
	const overrides = new Map<InstanceCell, number>();
	const local = new Map<string, number>();
	const formulas = new Map<string, string>();
	const expectedLocPin = new Map<string, number>();
	for (const { name, value, formula } of pose.cells) {
		const item = cell(name);
		if (!item) fail('UNSUPPORTED_GROUP_EDIT', `A stencil member has no ${name}.`);
		if (name.startsWith('LocPin') && effectiveFormula(item!)) {
			// The master places the pin with its own formula: it must follow the new size.
			expectedLocPin.set(name, value);
			continue;
		}
		assertOverridable(item, `The ${name}`);
		overrides.set(item!, value);
		local.set(name, value);
		formulas.set(name, formula);
	}
	const pins = new Map<InstanceCell, number>(
		(['Width', 'Height'] as const).flatMap((name) => {
			const item = cell(name);
			return item ? [[item, name === 'Width' ? pose.width : pose.height] as const] : [];
		}),
	);
	const { writes, pinned } = recalculateInstanceCaches(inst.scope, overrides, check, pins);
	for (const [item, value] of pinned) {
		assertOverridable(item, 'The size');
		local.set(item.name, value);
	}
	for (const [name, expected] of expectedLocPin) {
		const written = writes.find((write) => write.cell === cell(name));
		if (!written || !close(written.value, expected))
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'A stencil member places its pin with a formula that does not follow the group scaling.',
			);
	}
	const changed = new Map<string, Set<string>>();
	const touch = (owner: InstanceSheet, names: readonly string[]) => {
		const id = attribute(owner.instance, 'ID') ?? '';
		const set = changed.get(id) ?? new Set<string>();
		for (const name of names) set.add(name.toLowerCase());
		changed.set(id, set);
	};
	touch(inst.sheet, [...local.keys()]);
	for (const write of writes) touch(write.sheet, write.cell.names);
	assertNoPageDependents(root, shape, changed, check);
	return () => writeInstanceResult(inst.sheet, cell, local, writes, check, formulas);
}
