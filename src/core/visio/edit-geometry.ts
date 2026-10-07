import { createRectangle, createEllipse, createLine } from './edit-shape-create';
import { attribute, children } from './sheet';
import { fail } from './package-common';
import {
	assertVisioShapeUnreferenced,
	recalculateVisioCells,
	type VisioCellKey,
} from './edit-recalculate';
import { geometryChangedCells, type VisioGeometryEdit } from './edit-commands';
import { emptyMasterMoveProof, type MasterMoveProof } from './edit-master-move';
import {
	moveLocalLine,
	moveLocalLineEndpoint,
	proveLocalLine,
	assertLineTranslation,
	sameLineCoordinate,
} from './edit-line-move';
import {
	numeric,
	editableCell,
	setCell,
	cells,
	isLineSheet,
	admitted,
	protectedShape,
	resizeGeometry,
} from './edit-geometry-admission';

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
	const lineEditShapes = new Set<Element>();
	let fixedLine: ReadonlyMap<string, number> | undefined;
	const add = (cell: string) => changed.push({ pageId: edit.pageId, shapeId: edit.shapeId, cell });
	if (edit.type === 'create-line') {
		const shape = createLine(root, document, edit);
		lineEditShapes.add(shape);
		fixedLine = proveLocalLine(shape);
		expected = {
			width: Math.hypot(edit.endX - edit.beginX, edit.endY - edit.beginY),
			height: 0,
			x: (edit.beginX + edit.endX) / 2,
			y: (edit.beginY + edit.endY) / 2,
		};
		for (const name of geometryChangedCells(edit)) add(name);
	} else if (edit.type === 'create-rectangle' || edit.type === 'create-ellipse') {
		if (edit.type === 'create-ellipse') createEllipse(root, document, edit);
		else createRectangle(root, document, edit);
		expected = { width: edit.width, height: edit.height, x: edit.x, y: edit.y };
		for (const name of ['Width', 'Height', 'PinX', 'PinY']) add(name);
	} else {
		const shape = admitted(
			root,
			edit.shapeId,
			edit.type === 'move-shape' ? masterMovePins : new Set(),
			edit.type === 'move-shape' ? masterDimensions : new Map(),
			edit.type === 'move-shape' ? 'move' : edit.type === 'delete-shape' ? 'delete' : 'resize',
		);
		const local = cells(shape);
		const lineMove = edit.type === 'move-shape' && isLineSheet(local);
		protectedShape(
			shape,
			document,
			edit.type === 'move-shape' ? masterMovePins : new Set(),
			lineMove
				? ['LockBegin', 'LockEnd']
				: edit.type === 'move-line-endpoint'
					? [edit.endpoint === 'begin' ? 'LockBegin' : 'LockEnd']
					: [],
		);
		if (edit.type !== 'delete-shape')
			for (const connections of children(root, 'Connects'))
				for (const connection of children(connections, 'Connect'))
					if (['FromSheet', 'ToSheet'].some((name) => attribute(connection, name) === edit.shapeId))
						fail(
							'UNSUPPORTED_GEOMETRY_EDIT',
							'Glued connections need routing and endpoint recalculation outside this subset.',
						);
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
			if (lineMove) {
				const translation = moveLocalLine(shape, edit.pageId, edit.shapeId, edit.x, edit.y);
				changed.push(...translation.changed);
				fixedLine = translation.fixed;
				lineEditShapes.add(shape);
			} else
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
		} else if (edit.type === 'move-line-endpoint') {
			if (!isLineSheet(local))
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Endpoint editing requires a local line.');
			const endpoint = moveLocalLineEndpoint(shape, edit);
			changed.push(...endpoint.changed);
			expected = endpoint.expected;
			fixedLine = endpoint.fixed;
			lineEditShapes.add(shape);
			if (endpoint.expected.width !== numeric(local.get('Width')))
				resizeGeometry(shape, roots, edit, check, lineEditShapes);
		} else {
			const width = numeric(local.get('Width')),
				height = numeric(local.get('Height'));
			const lineResize = isLineSheet(local);
			if (lineResize) {
				if (edit.height !== 0)
					fail('UNSUPPORTED_GEOMETRY_EDIT', 'Line Width-cell resizing requires zero Height.');
				const proof = proveLocalLine(shape);
				fixedLine = new Map(
					[...proof].filter(([name]) => !['Width', 'LocPinX', 'LocPinY'].includes(name)),
				);
				lineEditShapes.add(shape);
			} else if (edit.height <= 0) fail('INVALID_EDIT', '2D resizing requires positive Height.');
			expected = {
				width: edit.width,
				height: edit.height,
				x: numeric(local.get('PinX'), width / 2),
				y: numeric(local.get('PinY'), height / 2),
			};
			if (width === edit.width && height === edit.height) return [];
			resizeGeometry(shape, roots, edit, check, lineEditShapes);
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
				if (!lineResize || name !== 'Width') editableCell(local.get(name));
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
	const affectedPages = recalculateVisioCells(roots, changed, {
		check,
		masterMovePins,
		lineEditShapes,
	});
	const resultShape = admitted(
		root,
		edit.shapeId,
		edit.type === 'move-shape' ? masterMovePins : new Set(),
		edit.type === 'move-shape' ? masterDimensions : new Map(),
		fixedLine ? (edit.type === 'resize-shape' ? 'resize' : 'move') : undefined,
	);
	if (fixedLine) assertLineTranslation(resultShape, fixedLine);
	if (edit.type === 'move-line-endpoint') proveLocalLine(resultShape);
	const result = cells(resultShape),
		provenResult = masterDimensions.get(resultShape);
	const equal = fixedLine ? sameLineCoordinate : (a: number, b: number) => a === b;
	if (
		expected &&
		(!equal(numeric(result.get('Width'), provenResult?.width), expected.width) ||
			!equal(numeric(result.get('Height'), provenResult?.height), expected.height) ||
			!equal(numeric(result.get('PinX'), expected.width / 2), expected.x) ||
			!equal(numeric(result.get('PinY'), expected.height / 2), expected.y))
	)
		fail(
			'EDIT_UNSUPPORTED_DEPENDENCY',
			'Dependent formulas would violate the requested geometry or fixed rotation pin.',
		);
	return [...new Set([edit.pageId, ...affectedPages])];
}
