import { attribute, children } from './sheet.js';
import { fail } from './package-common.js';
import {
	assertVisioShapeUnreferenced,
	recalculateVisioCells,
	type VisioCellKey,
} from './edit-recalculate.js';
import type { VisioGeometryEdit } from './edit-commands.js';
import { emptyMasterMoveProof, type MasterMoveProof } from './edit-master-move.js';
import {
	numeric,
	editableCell,
	setCell,
	cells,
	admitted,
	protectedShape,
	resizeGeometry,
} from './edit-geometry-admission.js';

function createRectangle(
	root: Element,
	edit: Extract<VisioGeometryEdit, { type: 'create-rectangle' }>,
): Element {
	const doc = root.ownerDocument!;
	const node = (name: string) => doc.createElementNS(root.namespaceURI, name);
	let container = children(root, 'Shapes')[0];
	if (children(root, 'Shapes').length > 1) fail('INVALID_SHAPE_ID', 'Duplicate Shapes containers.');
	const pending: Element[] = [root];
	while (pending.length) {
		const parent = pending.pop()!;
		for (const shapes of children(parent, 'Shapes'))
			for (const shape of children(shapes, 'Shape')) {
				if (attribute(shape, 'ID') === edit.shapeId)
					fail('INVALID_SHAPE_ID', 'Shape ID already exists.');
				pending.push(shape);
			}
	}
	if (!container) {
		container = node('Shapes');
		root.insertBefore(container, children(root, 'Connects')[0] ?? null);
	}
	const shape = node('Shape');
	shape.setAttribute('ID', edit.shapeId);
	shape.setAttribute('Type', 'Shape');
	for (const [name, value] of Object.entries({
		PinX: edit.x,
		PinY: edit.y,
		Width: edit.width,
		Height: edit.height,
		LocPinX: edit.width / 2,
		LocPinY: edit.height / 2,
		Angle: 0,
	}))
		setCell(
			shape,
			name,
			value,
			name === 'LocPinX' ? 'Width*0.5' : name === 'LocPinY' ? 'Height*0.5' : undefined,
		);
	const section = node('Section');
	section.setAttribute('N', 'Geometry');
	section.setAttribute('IX', '0');
	for (const [index, [x, y]] of [
		[0, 0],
		[1, 0],
		[1, 1],
		[0, 1],
		[0, 0],
	].entries()) {
		const row = node('Row');
		row.setAttribute('IX', String(index + 1));
		row.setAttribute('T', index ? 'RelLineTo' : 'RelMoveTo');
		setCell(row, 'X', x!);
		setCell(row, 'Y', y!);
		section.appendChild(row);
	}
	shape.appendChild(section);
	const text = node('Text');
	text.appendChild(doc.createTextNode(edit.text ?? ''));
	shape.appendChild(text);
	container.appendChild(shape);
	return shape;
}

export function applyGeometryEdit(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioGeometryEdit,
	check: () => void,
	masterMoveProof: MasterMoveProof = emptyMasterMoveProof(),
): readonly string[] {
	const masterMovePins = masterMoveProof.pins,
		masterDimensions = masterMoveProof.dimensions;
	check();
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const changed: VisioCellKey[] = [];
	let expected: { width: number; height: number; x: number; y: number } | undefined;
	const add = (cell: string) => changed.push({ pageId: edit.pageId, shapeId: edit.shapeId, cell });
	if (edit.type === 'create-rectangle') {
		createRectangle(root, edit);
		expected = { width: edit.width, height: edit.height, x: edit.x, y: edit.y };
		for (const name of ['Width', 'Height', 'PinX', 'PinY']) add(name);
	} else {
		const shape = admitted(
			root,
			edit.shapeId,
			edit.type === 'move-shape' ? masterMovePins : new Set(),
			edit.type === 'move-shape' ? masterDimensions : new Map(),
		);
		protectedShape(shape, document, edit.type === 'move-shape' ? masterMovePins : new Set());
		if (edit.type !== 'delete-shape')
			for (const connections of children(root, 'Connects'))
				for (const connection of children(connections, 'Connect'))
					if (['FromSheet', 'ToSheet'].some((name) => attribute(connection, name) === edit.shapeId))
						fail(
							'UNSUPPORTED_GEOMETRY_EDIT',
							'Glued connections need routing and endpoint recalculation outside this subset.',
						);
		const local = cells(shape);
		const unlocked = (
			name: 'LockMoveX' | 'LockMoveY' | 'LockWidth' | 'LockHeight' | 'LockAspect' | 'LockDelete',
		) => {
			if (numeric(local.get(name), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `${name} prevents this operation.`);
		};
		if (edit.type === 'delete-shape') {
			unlocked('LockDelete');
			assertVisioShapeUnreferenced(roots, edit.pageId, edit.shapeId, { check });
			shape.parentNode!.removeChild(shape);
			return [edit.pageId];
		}
		if (edit.type === 'move-shape') {
			const proven = masterDimensions.get(shape);
			expected = {
				width: numeric(local.get('Width'), proven?.width),
				height: numeric(local.get('Height'), proven?.height),
				x: edit.x,
				y: edit.y,
			};
			for (const [name, value, lock] of [
				['PinX', edit.x, 'LockMoveX'],
				['PinY', edit.y, 'LockMoveY'],
			] as const) {
				if (
					numeric(local.get(name), name === 'PinX' ? expected.width / 2 : expected.height / 2) ===
					value
				)
					continue;
				unlocked(lock);
				editableCell(local.get(name));
				setCell(shape, name, value);
				add(name);
			}
		} else {
			const width = numeric(local.get('Width')),
				height = numeric(local.get('Height'));
			expected = {
				width: edit.width,
				height: edit.height,
				x: numeric(local.get('PinX'), width / 2),
				y: numeric(local.get('PinY'), height / 2),
			};
			if (width === edit.width && height === edit.height) return [];
			resizeGeometry(shape, roots, edit, check);
			if (
				numeric(local.get('LockAspect'), 0) !== 0 &&
				Math.abs(edit.width / width - edit.height / height) > 1e-9
			)
				fail('EDIT_PROTECTED_CELL', 'LockAspect prevents changing the aspect ratio.');
			// Make implicit pins explicit before dimensions change. Resizing holds the rotation pin fixed.
			for (const [name, value] of [
				['PinX', width / 2],
				['PinY', height / 2],
			] as const)
				if (!local.has(name)) setCell(shape, name, value);
			for (const [name, value, lock] of [
				['Width', edit.width, 'LockWidth'],
				['Height', edit.height, 'LockHeight'],
			] as const) {
				if (numeric(local.get(name)) === value) continue;
				unlocked(lock);
				editableCell(local.get(name));
				setCell(shape, name, value);
				add(name);
			}
			for (const [name, dimension, value] of [
				['LocPinX', 'Width', edit.width / 2],
				['LocPinY', 'Height', edit.height / 2],
			] as const)
				if (!local.has(name)) setCell(shape, name, value, `${dimension}*0.5`);
		}
	}
	if (!changed.length) return [];
	const affectedPages = recalculateVisioCells(roots, changed, { check, masterMovePins });
	const resultShape = admitted(
		root,
		edit.shapeId,
		edit.type === 'move-shape' ? masterMovePins : new Set(),
		edit.type === 'move-shape' ? masterDimensions : new Map(),
	);
	const result = cells(resultShape),
		provenResult = masterDimensions.get(resultShape);
	if (
		expected &&
		(numeric(result.get('Width'), provenResult?.width) !== expected.width ||
			numeric(result.get('Height'), provenResult?.height) !== expected.height ||
			numeric(result.get('PinX'), expected.width / 2) !== expected.x ||
			numeric(result.get('PinY'), expected.height / 2) !== expected.y)
	)
		fail(
			'EDIT_UNSUPPORTED_DEPENDENCY',
			'Dependent formulas would violate the requested geometry or fixed rotation pin.',
		);
	return [...new Set([edit.pageId, ...affectedPages])];
}
