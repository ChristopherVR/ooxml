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
	const members: Element[] = [],
		nested = new Set<Element>([group]);
	const pending = [group];
	while (pending.length) {
		check();
		const parent = pending.pop()!;
		const direct = children(children(parent, 'Shapes')[0], 'Shape');
		if (!direct.length || members.length + direct.length > 500 || nested.size > 64)
			fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resizing requires a bounded local group tree.');
		for (const shape of direct) {
			members.push(shape);
			const type = attribute(shape, 'Type');
			if (type === 'Group') {
				nested.add(shape);
				pending.push(shape);
			} else if (type !== 'Shape' || children(shape, 'Shapes').length)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resizing requires groups and 2D shapes.');
		}
	}
	if (
		[group, ...members].some(
			(shape) => shape.hasAttribute('Master') || shape.hasAttribute('MasterShape'),
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Group resizing requires a local group of 2D shapes.');
	const ids = proveLocalShapeTree(group, check);
	assertUngluedTree(root, ids, 'Group resizing cannot update glued members or group connections.');
	const shapes = new Set([group, ...members]);
	for (const shape of shapes) {
		check();
		assertCanonicalResizeCells(shape);
		protectedShape(shape, document);
		assertUnlayeredShape(shape, document);
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
	return { group, members, shapes, groups: nested };
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
	/** Scale the direct members of `parent` for its resize to (newW, newH), then recurse into
	 * nested groups with their own resulting size. */
	const plan = (parent: Element, newW: number, newH: number, parentMode: number): void => {
		const parentId = attribute(parent, 'ID')!;
		const width = read(parent, 'Width'),
			height = read(parent, 'Height');
		const sx = newW / width,
			sy = newH / height;
		for (const shape of children(children(parent, 'Shapes')[0], 'Shape')) {
			check();
			const id = attribute(shape, 'ID')!,
				local = cells(shape),
				scaling = mode(shape, parentMode) === 2;
			const childW = read(shape, 'Width'),
				childH = read(shape, 'Height');
			// Visio scales a rotated member's local size by the squared direction cosines of its
			// axes (verified against Visio 16): W' = W(cos^2 sx + sin^2 sy), H' = H(sin^2 sx + cos^2 sy).
			const angle = read(shape, 'Angle', 0);
			const c2 = Math.cos(angle) ** 2,
				s2 = Math.sin(angle) ** 2;
			const sizeX = scaling ? c2 * sx + s2 * sy : 1,
				sizeY = scaling ? s2 * sx + c2 * sy : 1;
			if (numeric(local.get('LockAspect'), 0) !== 0 && !resizeClose(sizeX, sizeY))
				fail('EDIT_PROTECTED_CELL', 'A group member LockAspect prevents resizing.');
			const isGroup = attribute(shape, 'Type') === 'Group';
			if (
				scaling &&
				(!resizeClose(sx, 1) || !resizeClose(sy, 1)) &&
				(!isGroup || children(shape, 'Section').some((item) => attribute(item, 'N') === 'Geometry'))
			)
				resizeGeometry(
					shape,
					roots,
					{ ...edit, shapeId: id },
					check,
					new Set(),
					new Set(),
					scope.shapes,
				);
			const ref = `Sheet.${parentId}!`;
			const sized = (own: 'Width' | 'Height', size: number) => {
				const terms: string[] = [];
				const [fx, fy] = own === 'Width' ? [c2, s2] : [s2, c2];
				if (fx > 1e-12) terms.push(`${ref}Width*${(size * fx) / width}`);
				if (fy > 1e-12) terms.push(`${ref}Height*${(size * fy) / height}`);
				return terms.join('+');
			};
			for (const [name, factor, formula, lock] of [
				['PinX', sx, `${ref}Width*${read(shape, 'PinX') / width}`, 'LockMoveX'],
				['PinY', sy, `${ref}Height*${read(shape, 'PinY') / height}`, 'LockMoveY'],
				['Width', sizeX, sized('Width', childW), 'LockWidth'],
				['Height', sizeY, sized('Height', childH), 'LockHeight'],
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
			if (isGroup) plan(shape, childW * sizeX, childH * sizeY, scaling ? mode(shape, 2) : 1);
		}
	};
	plan(scope.group, edit.width, edit.height, mode(scope.group, 2));
	return { writes, expected };
}
