import { attribute, children } from './sheet';
import { fail } from './package-common';
import {
	cells,
	numeric,
	editableCell,
	protectedShape,
	resizeGeometry,
	guardedCell,
} from './edit-geometry-admission';
import { assertUnlayeredShape } from './edit-style-admission';
import { proveLocalShapeTree, assertUngluedTree } from './edit-group-rotation';
import { analyzeVisioFormula, visioFormulaCachedValue, type VisioFormulaValue } from './formula';
import { executableCellFormula } from './cell-formula';
import { key } from './edit-recalculate-index';
import { assertCanonicalResizeCells, resizeClose, type VisioResizeWrite } from './edit-resize-plan';
import type { VisioGeometryEdit } from './edit-commands';

type Resize = Extract<VisioGeometryEdit, { type: 'resize-shape' }>;
type Evaluate = (id: string) => VisioFormulaValue;
export interface LocalGroupResize {
	group: Element;
	members: Element[];
	shapes: ReadonlySet<Element>;
	groups: ReadonlySet<Element>;
}

/** Coarse source proof only. Formula caches, dimensions and locks are checked by the plan. */
export function localGroupResize(
	root: Element,
	id: string,
	document: Element,
	check: () => void,
): LocalGroupResize | undefined {
	const candidates = children(children(root, 'Shapes')[0], 'Shape').filter(
		(shape) => attribute(shape, 'ID') === id,
	);
	if (candidates.length !== 1 || attribute(candidates[0], 'Type') !== 'Group') return undefined;
	const group = candidates[0]!;
	const members = children(children(group, 'Shapes')[0], 'Shape');
	if (
		[group, ...members].some(
			(shape) => shape.hasAttribute('Master') || shape.hasAttribute('MasterShape'),
		) ||
		!members.length ||
		members.length > 500 ||
		members.some(
			(shape) => attribute(shape, 'Type') !== 'Shape' || children(shape, 'Shapes').length,
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resizing requires a flat local group of 2D shapes.');
	const ids = proveLocalShapeTree(group, check);
	assertUngluedTree(root, ids, 'Group resizing cannot update glued members or group connections.');
	const shapes = new Set([group, ...members]);
	for (const shape of shapes) {
		check();
		assertCanonicalResizeCells(shape);
		protectedShape(shape, document);
		assertUnlayeredShape(shape, document);
		if (shape !== group && Math.abs(Math.sin(numeric(cells(shape).get('Angle'), 0))) > 1e-10)
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resizing requires axis-aligned local members.');
		// Text-measured sizes are refreshed after a transaction, outside this resize proof.
		for (const node of Array.from(shape.childNodes).filter(
			(node): node is Element => node.nodeType === 1 && (node as Element).localName !== 'Shapes',
		)) {
			for (const cell of [node, ...Array.from(node.getElementsByTagName('*'))]) {
				const formula = executableCellFormula(attribute(cell, 'F'));
				if (formula && analyzeVisioFormula(formula).unsupportedFunctions.length)
					fail(
						'EDIT_UNSUPPORTED_FORMULA',
						'Group resize cannot refresh unsupported or text-measured formulas.',
					);
			}
		}
	}
	return { group, members, shapes, groups: new Set([group]) };
}

/** Member scaling is saved as ordinary ShapeSheet dependencies on its parent. Existing guarded
 * formulas remain intact and must evaluate to the requested pose. ResizeMode 0 inherits the
 * group's mode; a top-level default group scales. Mode 1 keeps member size, mode 2 scales it.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/resizemode-cell-shape-transform-section
 */
export function planGroupResize(
	scope: LocalGroupResize,
	roots: ReadonlyMap<string, Element>,
	edit: Resize,
	source: Evaluate,
	check: () => void,
) {
	const writes: VisioResizeWrite[] = [];
	const expected = new Map<string, number>();
	const read = (shape: Element, name: string, fallback?: number) => {
		const local = cells(shape);
		return local.has(name) || ['LocPinX', 'LocPinY'].includes(name)
			? source(key({ pageId: edit.pageId, shapeId: attribute(shape, 'ID')!, cell: name })).value
			: numeric(local.get(name), fallback);
	};
	const mode = (shape: Element, fallback: number) => {
		const cell = cells(shape).get('ResizeMode');
		if (!cell) return fallback;
		const formula = executableCellFormula(attribute(cell, 'F'));
		if (
			attribute(cell, 'F') === 'Inh' ||
			(formula && analyzeVisioFormula(formula).references.length) ||
			visioFormulaCachedValue('0', attribute(cell, 'U')).unit !== 'scalar'
		)
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resize needs a constant local ResizeMode.');
		const value = read(shape, 'ResizeMode');
		if (![0, 1, 2].includes(value))
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Unsupported group ResizeMode.');
		return value || fallback;
	};
	const width = read(scope.group, 'Width'),
		height = read(scope.group, 'Height');
	const sx = edit.width / width,
		sy = edit.height / height,
		groupMode = mode(scope.group, 2);
	for (const shape of scope.members) {
		check();
		const id = attribute(shape, 'ID')!,
			local = cells(shape),
			scaling = mode(shape, groupMode) === 2;
		const childW = read(shape, 'Width'),
			childH = read(shape, 'Height');
		const sizeX = scaling ? sx : 1,
			sizeY = scaling ? sy : 1;
		if (numeric(local.get('LockAspect'), 0) !== 0 && !resizeClose(sizeX, sizeY))
			fail('EDIT_PROTECTED_CELL', 'A group member LockAspect prevents resizing.');
		if (scaling && (!resizeClose(sx, 1) || !resizeClose(sy, 1)))
			resizeGeometry(
				shape,
				roots,
				{ ...edit, shapeId: id },
				check,
				new Set(),
				new Set(),
				scope.shapes,
			);
		for (const [name, factor, formula, lock] of [
			['PinX', sx, `Sheet.${edit.shapeId}!Width*${read(shape, 'PinX') / width}`, 'LockMoveX'],
			['PinY', sy, `Sheet.${edit.shapeId}!Height*${read(shape, 'PinY') / height}`, 'LockMoveY'],
			['Width', sizeX, `Sheet.${edit.shapeId}!Width*${childW / width}`, 'LockWidth'],
			['Height', sizeY, `Sheet.${edit.shapeId}!Height*${childH / height}`, 'LockHeight'],
			['LocPinX', sizeX, `Width*${read(shape, 'LocPinX') / childW}`, ''],
			['LocPinY', sizeY, `Height*${read(shape, 'LocPinY') / childH}`, ''],
		] as const) {
			const old = read(shape, name),
				value = old * factor;
			if (
				!Number.isFinite(value) ||
				Math.abs(value) > 1e6 ||
				(['Width', 'Height'].includes(name) && value <= 0)
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Resized group member exceeds coordinate limits.');
			expected.set(key({ pageId: edit.pageId, shapeId: id, cell: name }), value);
			if (resizeClose(value, old)) continue;
			if (lock && numeric(local.get(lock), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `A group member ${lock} prevents resizing.`);
			const node = local.get(name);
			const existing = executableCellFormula(attribute(node, 'F'));
			// Preserve authored relationships. The final projection must prove that they already
			// follow the requested member pose; otherwise refuse instead of severing a dependency.
			if (guardedCell(node) || (existing && analyzeVisioFormula(existing).references.length))
				continue;
			editableCell(node, (reference) =>
				source(
					key({ pageId: edit.pageId, shapeId: reference.shapeId ?? id, cell: reference.cell }),
				),
			);
			writes.push({ shape, name, value, formula });
		}
		for (const name of ['Angle', 'FlipX', 'FlipY'])
			expected.set(key({ pageId: edit.pageId, shapeId: id, cell: name }), read(shape, name, 0));
	}
	return { writes, expected };
}
