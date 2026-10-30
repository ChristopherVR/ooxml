import type { VisioGeometryEdit } from './edit-commands';
import {
	admitted,
	cells,
	numeric,
	protectedShape,
	resizeGeometry,
	editableCell,
	guardedCell,
	setCell,
	isLineSheet,
} from './edit-geometry-admission';
import { indexCells, key, type IndexedCell } from './edit-recalculate-index';
import { createVisioCellEvaluator } from './edit-recalculate-values';
import { parseVisioFormula, analyzeVisioFormula, visioFormulaCachedValue } from './formula';
import { transform } from './geometry';
import { attribute, children, VISIO_NS, VISIO_LEGACY_NS } from './sheet';
import { visioAnchoredResizeGeometry } from './resize-anchor';
import { fail } from './package-common';
import { localGroupResize, planGroupResize } from './edit-group-resize';
import {
	assertCanonicalResizeCells,
	resizeClosure as closure,
	resizeClose as close,
	type VisioResizeWrite as Write,
} from './edit-resize-plan';

type Resize = Extract<VisioGeometryEdit, { type: 'resize-shape' }>;
/** All source and proposed caches are evaluated before writing either geometry or dependent caches. */
export function resizeVisioShapeAtAnchor(
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: Resize,
	check: () => void,
	glueShapes: ReadonlySet<Element> = new Set(),
): readonly string[] {
	check();
	const root = roots.get(edit.pageId);
	if (!root) fail('EDIT_TARGET_NOT_FOUND', 'Page does not exist.');
	const group = localGroupResize(root, edit.shapeId, document, check);
	const shape = admitted(root, edit.shapeId, new Set(), new Map(), undefined, group?.groups);
	const local = cells(shape);
	assertCanonicalResizeCells(shape);
	if (!edit.anchor || isLineSheet(cells(shape)) || edit.height <= 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Anchored resizing requires an ordinary local 2D shape.');
	if (
		!group &&
		Array.from(shape.getElementsByTagName('*')).some(
			(node) =>
				node.localName === 'Shape' && [VISIO_NS, VISIO_LEGACY_NS].includes(node.namespaceURI ?? ''),
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Nested shape identities cannot be resized as a local leaf.');
	protectedShape(shape, document);
	// Layer membership does not touch geometry; edit.ts refuses a shape on a locked layer.
	for (const container of children(root, 'Connects'))
		for (const connection of children(container, 'Connect'))
			if (
				!glueShapes.has(shape) &&
				['FromSheet', 'ToSheet'].some((name) => attribute(connection, name) === edit.shapeId)
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Glued shapes require routing outside this subset.');
	const graph = indexCells(roots, {
		check,
		scope: 'affected',
		glueShapes,
		...(group ? { groupResizeShapes: group.shapes } : {}),
	});
	const id = (cell: string) => key({ pageId: edit.pageId, shapeId: edit.shapeId, cell });
	const source = createVisioCellEvaluator(graph, { check, glueShapes }, true);
	const read = (name: string, fallback?: number): number =>
		graph.has(id(name)) ? source(id(name)).value : numeric(local.get(name), fallback);
	const width = read('Width'),
		height = read('Height');
	const pinX = read('PinX', width / 2),
		pinY = read('PinY', height / 2);
	const locX = read('LocPinX', width / 2),
		locY = read('LocPinY', height / 2);
	const angle = read('Angle', 0),
		flipX = read('FlipX', 0),
		flipY = read('FlipY', 0);
	for (const name of ['Angle', 'FlipX', 'FlipY']) {
		const unit = attribute(local.get(name), 'U');
		if (unit && visioFormulaCachedValue('0', unit).unit !== (name === 'Angle' ? 'angle' : 'scalar'))
			fail('EDIT_FORMULA_UNIT', 'Rotation and flip cells require their declared dimensions.');
	}
	if (![0, 1].includes(flipX) || ![0, 1].includes(flipY))
		fail('EDIT_FORMULA_UNIT', 'Flip cells must be scalar booleans.');
	if (width === edit.width && height === edit.height) return [];
	if (
		!group ||
		children(shape, 'Section').some((section) => attribute(section, 'N') === 'Geometry')
	)
		resizeGeometry(shape, roots, edit, check, new Set(), glueShapes, group?.shapes);
	if (numeric(local.get('LockAspect'), 0) !== 0 && !close(edit.width / width, edit.height / height))
		fail('EDIT_PROTECTED_CELL', 'LockAspect prevents changing the aspect ratio.');
	const geometry = visioAnchoredResizeGeometry(
		{
			width,
			height,
			pinX,
			pinY,
			transform: transform(pinX, pinY, locX, locY, angle, flipX === 1, flipY === 1),
		},
		edit,
		edit.anchor,
	);
	if (
		!geometry ||
		[geometry.pinX, geometry.pinY, geometry.locPinX, geometry.locPinY].some(
			(value) => Math.abs(value) > 1e6,
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Anchored geometry exceeds coordinate limits.');
	const writes: Write[] = [];
	const writable = (name: string) =>
		editableCell(local.get(name), (reference) =>
			source(
				key({
					pageId: edit.pageId,
					shapeId: reference.shapeId ?? edit.shapeId,
					cell: reference.cell,
				}),
			),
		);
	const unlocked = (name: string) => {
		if (numeric(local.get(name), 0) !== 0)
			fail('EDIT_PROTECTED_CELL', `${name} prevents this operation.`);
	};
	for (const [name, value, old] of [
		['Width', edit.width, width],
		['Height', edit.height, height],
	] as const)
		if (value !== old) {
			unlocked(`Lock${name}`);
			writable(name);
			writes.push({ name, value });
		}
	const changedDimensions = writes.map((write) => id(write.name));
	const dimensionClosure = closure(graph, changedDimensions, check);
	for (const [name, value, old] of [
		['PinX', geometry.pinX, pinX],
		['PinY', geometry.pinY, pinY],
	] as const)
		if (
			!local.has(name) ||
			!close(value, old) ||
			(dimensionClosure.has(id(name)) && !guardedCell(local.get(name)))
		) {
			if (!close(value, old)) unlocked(name === 'PinX' ? 'LockMoveX' : 'LockMoveY');
			writable(name);
			writes.push({ name, value: close(value, old) ? old : value });
		}
	const projected = new Map(graph);
	const writeId = (write: Write) =>
		key({ pageId: edit.pageId, shapeId: attribute(write.shape ?? shape, 'ID')!, cell: write.name });
	const install = (write: Write): void => {
		const owner = write.shape ?? shape,
			ownerId = attribute(owner, 'ID')!;
		const original = cells(owner).get(write.name);
		const node = original
			? (original.cloneNode(true) as Element)
			: shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Cell');
		node.setAttribute('N', write.name);
		node.setAttribute('V', String(write.value));
		if (write.formula) node.setAttribute('F', write.formula);
		else node.removeAttribute('F');
		const cell: IndexedCell = {
			pageId: edit.pageId,
			shapeId: ownerId,
			cell: write.name,
			node,
			unsafe: false,
			unit: 'length',
			dependencies: [],
		};
		if (write.formula) {
			cell.formula = parseVisioFormula(write.formula);
			cell.dependencies = analyzeVisioFormula(cell.formula).references.map((reference) =>
				key({ pageId: edit.pageId, shapeId: reference.shapeId ?? ownerId, cell: reference.cell }),
			);
		}
		projected.set(writeId(write), cell);
	};
	for (const write of writes) install(write);
	const guardedPins: { name: string; value: number }[] = [];
	for (const [name, dimension, old, oldDimension, next] of [
		['LocPinX', 'Width', locX, width, geometry.locPinX],
		['LocPinY', 'Height', locY, height, geometry.locPinY],
	] as const) {
		if ((dimension === 'Width' ? edit.width : edit.height) === oldDimension) continue;
		if (guardedCell(local.get(name))) {
			guardedPins.push({ name, value: next });
		} else {
			writable(name);
			const write = { name, value: next, formula: `${dimension}*${old / oldDimension}` };
			writes.push(write);
			install(write);
		}
	}
	const projectedPins = createVisioCellEvaluator(projected, { check });
	for (const pin of guardedPins)
		if (!close(projectedPins(id(pin.name)).value, pin.value))
			fail(
				'EDIT_PROTECTED_CELL',
				'Guarded local pin does not scale proportionally with the requested dimension.',
			);
	const members = group && planGroupResize(group, roots, edit, source, check);
	for (const write of members?.writes ?? []) {
		writes.push(write);
		install(write);
	}
	const affected = closure(projected, writes.map(writeId), check);
	for (const name of ['Angle', 'FlipX', 'FlipY'])
		if (affected.has(id(name)))
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Anchored resizing cannot change dependent Angle or Flip formulas.',
			);
	for (const cellId of affected)
		if (/^Field\./i.test(projected.get(cellId)!.cell))
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Affected text fields need native display-cache regeneration.',
			);
	const evaluate = createVisioCellEvaluator(projected, { check }),
		values = new Map<string, number>();
	for (const cellId of affected) {
		if (graph.has(cellId)) source(cellId);
		values.set(cellId, evaluate(cellId).value);
	}
	for (const [name, expected] of [
		['Width', edit.width],
		['Height', edit.height],
		['PinX', geometry.pinX],
		['PinY', geometry.pinY],
		['LocPinX', geometry.locPinX],
		['LocPinY', geometry.locPinY],
	] as const)
		if (!close(projected.has(id(name)) ? evaluate(id(name)).value : read(name), expected))
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'Dependent formulas would violate the requested anchor or dimensions.',
			);
	for (const [cellId, expected] of members?.expected ?? [])
		if (!close(projected.has(cellId) ? evaluate(cellId).value : 0, expected))
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'Dependent formulas would violate a resized group member pose.',
			);
	check();
	for (const write of writes) setCell(write.shape ?? shape, write.name, write.value, write.formula);
	const dirty = new Set([edit.pageId]);
	for (const [cellId, value] of values) {
		const item = graph.get(cellId);
		if (!item?.node || !item.formula || writes.some((write) => writeId(write) === cellId)) continue;
		if (attribute(item.node, 'V') !== String(value)) {
			item.node.setAttribute('V', String(value));
			dirty.add(item.pageId);
		}
	}
	return [...dirty];
}
