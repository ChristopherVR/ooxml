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
import { parseVisioFormula, visioFormulaCachedValue } from './formula';
import { transform } from './geometry';
import { attribute, children, VISIO_NS, VISIO_LEGACY_NS } from './sheet';
import { assertUnlayeredShape } from './edit-style-admission';
import { visioAnchoredResizeGeometry } from './resize-anchor';
import { fail } from './package-common';

type Resize = Extract<VisioGeometryEdit, { type: 'resize-shape' }>;
interface Write {
	name: string;
	value: number;
	formula?: string;
}
const close = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-10 * Math.max(1, Math.abs(a), Math.abs(b));

function closure(
	graph: ReadonlyMap<string, IndexedCell>,
	changed: readonly string[],
	check: () => void,
) {
	const reverse = new Map<string, string[]>();
	for (const [id, item] of graph) {
		check();
		for (const dependency of item.dependencies) {
			const list = reverse.get(dependency) ?? [];
			list.push(id);
			reverse.set(dependency, list);
		}
	}
	const affected = new Set(changed),
		pending = [...changed];
	for (let index = 0; index < pending.length; index++) {
		check();
		for (const id of reverse.get(pending[index]!) ?? [])
			if (!affected.has(id)) {
				affected.add(id);
				pending.push(id);
			}
		if (affected.size > 10_000)
			fail('LIMIT_FORMULA_AFFECTED', 'Affected ShapeSheet cell limit exceeded.');
	}
	return affected;
}
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
	const shape = admitted(root, edit.shapeId);
	const local = cells(shape);
	const canonical = [
		'Width',
		'Height',
		'PinX',
		'PinY',
		'LocPinX',
		'LocPinY',
		'Angle',
		'FlipX',
		'FlipY',
		'LockMoveX',
		'LockMoveY',
		'LockWidth',
		'LockHeight',
		'LockAspect',
		'LockDelete',
	];
	for (const name of local.keys())
		if (
			canonical.some(
				(expected) => expected.toLowerCase() === name.toLowerCase() && expected !== name,
			)
		)
			fail(
				'EDIT_AMBIGUOUS_CELL',
				'Noncanonical transform and protection cell names are unsupported.',
			);
	if (!edit.anchor || isLineSheet(cells(shape)) || edit.height <= 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Anchored resizing requires an ordinary local 2D shape.');
	if (
		Array.from(shape.getElementsByTagName('*')).some(
			(node) =>
				node.localName === 'Shape' && [VISIO_NS, VISIO_LEGACY_NS].includes(node.namespaceURI ?? ''),
		)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Nested shape identities cannot be resized as a local leaf.');
	protectedShape(shape, document);
	assertUnlayeredShape(shape, document);
	for (const container of children(root, 'Connects'))
		for (const connection of children(container, 'Connect'))
			if (
				!glueShapes.has(shape) &&
				['FromSheet', 'ToSheet'].some((name) => attribute(connection, name) === edit.shapeId)
			)
				fail('UNSUPPORTED_GEOMETRY_EDIT', 'Glued shapes require routing outside this subset.');
	const graph = indexCells(roots, { check, glueShapes });
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
	resizeGeometry(shape, roots, edit, check);
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
	const install = (write: Write): void => {
		const original = local.get(write.name);
		const node = original
			? (original.cloneNode(true) as Element)
			: shape.ownerDocument!.createElementNS(shape.namespaceURI, 'Cell');
		node.setAttribute('N', write.name);
		node.setAttribute('V', String(write.value));
		if (write.formula) node.setAttribute('F', write.formula);
		else node.removeAttribute('F');
		const cell: IndexedCell = {
			pageId: edit.pageId,
			shapeId: edit.shapeId,
			cell: write.name,
			node,
			unsafe: false,
			unit: 'length',
			dependencies: [],
		};
		if (write.formula) {
			cell.formula = parseVisioFormula(write.formula);
			cell.dependencies = [id(write.name === 'LocPinX' ? 'Width' : 'Height')];
		}
		projected.set(id(write.name), cell);
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
	const affected = closure(
		projected,
		writes.map((write) => id(write.name)),
		check,
	);
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
	check();
	for (const write of writes) setCell(shape, write.name, write.value, write.formula);
	const dirty = new Set([edit.pageId]);
	for (const [cellId, value] of values) {
		const item = graph.get(cellId);
		if (!item?.node || !item.formula || writes.some((write) => id(write.name) === cellId)) continue;
		if (attribute(item.node, 'V') !== String(value)) {
			item.node.setAttribute('V', String(value));
			dirty.add(item.pageId);
		}
	}
	return [...dirty];
}
