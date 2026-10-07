import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, editableCell, numeric, setCell } from './edit-geometry-admission';
import { executableCellFormula } from './cell-formula';
import type { VisioCellKey } from './edit-recalculate';
import type { VisioGeometryEdit } from './edit-commands';
import { visioFormulaCachedValue, evaluateVisioFormula } from './formula';

export const sameLineCoordinate = (a: number, b: number) =>
	Math.abs(a - b) <= 1e-12 * Math.max(1, Math.abs(a), Math.abs(b));

function lengthUnit(node: Element | undefined): void {
	const unit = attribute(node, 'U');
	if (unit && visioFormulaCachedValue('0', unit).unit !== 'length')
		fail('EDIT_FORMULA_UNIT', 'Line coordinates require length units.');
}
function staticLength(node: Element | undefined): void {
	editableCell(node);
	lengthUnit(node);
	const source = executableCellFormula(attribute(node, 'F'));
	if (!source) return;
	const result = evaluateVisioFormula(source, () =>
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Unexpected constant dependency.'),
	);
	if (result.unit !== 'scalar' && result.unit !== 'length')
		fail('EDIT_FORMULA_UNIT', 'A line length formula has incompatible dimensions.');
	if (!sameLineCoordinate(result.value, numeric(node)))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'A constant line length cache is stale.');
}

/** Shared native DrawLine admission, including an explicit Width-cell override. */
export function proveLocalLine(shape: Element): ReadonlyMap<string, number> {
	const local = cells(shape);
	const value = (name: string) => numeric(local.get(name));
	const canonical = (name: string, expected: string) => {
		const source = executableCellFormula(attribute(local.get(name), 'F'));
		if (source?.replace(/\s+/g, '').toLowerCase() !== expected.toLowerCase())
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Line midpoint and length formulas must be locally proven.',
			);
	};
	canonical('PinX', '(BeginX+EndX)/2');
	canonical('PinY', '(BeginY+EndY)/2');
	const widthFormula = executableCellFormula(attribute(local.get('Width'), 'F'));
	const endpointLength =
		widthFormula?.replace(/\s+/g, '').toLowerCase() === 'sqrt((endx-beginx)^2+(endy-beginy)^2)';
	if (!endpointLength) {
		staticLength(local.get('Width'));
	}
	const bx = value('BeginX'),
		by = value('BeginY'),
		ex = value('EndX'),
		ey = value('EndY');
	for (const name of ['BeginX', 'BeginY', 'EndX', 'EndY']) {
		lengthUnit(local.get(name));
	}
	const width = value('Width');
	value('Angle');
	const angleUnit = attribute(local.get('Angle'), 'U');
	if (angleUnit && visioFormulaCachedValue('0', angleUnit).unit !== 'angle')
		fail('EDIT_FORMULA_UNIT', 'Line rotation must use angle units.');
	if (
		!(width > 0) ||
		value('Height') !== 0 ||
		(endpointLength && !sameLineCoordinate(width, Math.hypot(ex - bx, ey - by))) ||
		!sameLineCoordinate(value('PinX'), (bx + ex) / 2) ||
		!sameLineCoordinate(value('PinY'), (by + ey) / 2)
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Line transform caches disagree with their endpoints.');
	const geometries = children(shape, 'Section').filter(
		(node) => attribute(node, 'N') === 'Geometry',
	);
	const rows = children(geometries[0], 'Row');
	if (
		geometries.length !== 1 ||
		rows.length !== 2 ||
		attribute(rows[0], 'T') !== 'MoveTo' ||
		attribute(rows[1], 'T') !== 'LineTo' ||
		[geometries[0]!, ...rows].some((node) =>
			['1', 'true'].includes(attribute(node, 'Del') ?? ''),
		) ||
		!rows.every((row, index) => {
			const coordinates = cells(row);
			for (const name of ['X', 'Y']) {
				const node = coordinates.get(name);
				lengthUnit(node);
				const formula = executableCellFormula(attribute(node, 'F'))
					?.replace(/\s+/g, '')
					.toLowerCase();
				if (name !== 'X' || formula !== `width*${index}`) staticLength(node);
			}
			return (
				sameLineCoordinate(numeric(coordinates.get('X')), index ? width : 0) &&
				numeric(coordinates.get('Y')) === 0
			);
		})
	)
		fail(
			'UNSUPPORTED_GEOMETRY_EDIT',
			'Only local straight-line geometry is proven for translation.',
		);
	return new Map(
		['Width', 'Height', 'Angle', 'LocPinX', 'LocPinY', 'FlipX', 'FlipY'].map(
			(name) => [name, numeric(local.get(name), name === 'LocPinX' ? width / 2 : 0)] as const,
		),
	);
}

/** Native DrawLine leaves: midpoint/length formulas survive endpoint translation. */
export function moveLocalLine(
	shape: Element,
	pageId: string,
	shapeId: string,
	x: number,
	y: number,
): { changed: VisioCellKey[]; fixed: ReadonlyMap<string, number> } {
	const fixed = proveLocalLine(shape);
	const local = cells(shape);
	const value = (name: string) => numeric(local.get(name));
	const changed: VisioCellKey[] = [];
	if (x !== value('PinX') || y !== value('PinY'))
		for (const name of ['LockBegin', 'LockEnd'])
			if (numeric(local.get(name), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `${name} prevents moving a locked endpoint.`);
	for (const [axis, target] of [
		['X', x],
		['Y', y],
	] as const) {
		const delta = target - value(`Pin${axis}`);
		if (delta === 0) continue;
		if (numeric(local.get(`LockMove${axis}`), 0) !== 0)
			fail('EDIT_PROTECTED_CELL', `LockMove${axis} prevents this operation.`);
		for (const prefix of ['Begin', 'End']) {
			const name = `${prefix}${axis}`;
			editableCell(local.get(name));
			setCell(shape, name, value(name) + delta);
			numeric(local.get(name));
			changed.push({ pageId, shapeId, cell: name });
		}
	}
	return { changed, fixed };
}

export function assertLineTranslation(shape: Element, fixed: ReadonlyMap<string, number>): void {
	const local = cells(shape);
	for (const [name, expected] of fixed)
		if (!sameLineCoordinate(numeric(local.get(name), expected), expected))
			fail(
				'EDIT_UNSUPPORTED_DEPENDENCY',
				'Line translation would alter length, angle or local transform.',
			);
}

/** Native endpoint-cell assignment, retaining the derived midpoint, length and angle. */
export function moveLocalLineEndpoint(
	shape: Element,
	edit: Extract<VisioGeometryEdit, { type: 'move-line-endpoint' }>,
) {
	const proof = proveLocalLine(shape);
	const local = cells(shape);
	const formula = (name: string) =>
		executableCellFormula(attribute(local.get(name), 'F'))
			?.replace(/\s+/g, '')
			.toLowerCase();
	for (const [name, required] of [
		['Width', 'sqrt((endx-beginx)^2+(endy-beginy)^2)'],
		['Angle', 'atan2(endy-beginy,endx-beginx)'],
		['LocPinX', 'width*0.5'],
		['LocPinY', 'height*0.5'],
	])
		if (formula(name!) !== required)
			fail(
				'UNSUPPORTED_GEOMETRY_EDIT',
				'Endpoint editing requires native derived transform formulas.',
			);
	if (numeric(local.get('FlipX'), 0) !== 0 || numeric(local.get('FlipY'), 0) !== 0)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Flipped endpoint transforms require native evidence.');
	if (
		!sameLineCoordinate(
			proof.get('Angle')!,
			Math.atan2(
				numeric(local.get('EndY')) - numeric(local.get('BeginY')),
				numeric(local.get('EndX')) - numeric(local.get('BeginX')),
			),
		) ||
		!sameLineCoordinate(proof.get('LocPinX')!, proof.get('Width')! / 2) ||
		proof.get('LocPinY') !== 0
	)
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Derived endpoint transform caches are stale.');
	const prefix = edit.endpoint === 'begin' ? 'Begin' : 'End';
	const opposite = edit.endpoint === 'begin' ? 'End' : 'Begin';
	const ox = numeric(local.get(`${opposite}X`)),
		oy = numeric(local.get(`${opposite}Y`));
	const dx = edit.endpoint === 'begin' ? ox - edit.x : edit.x - ox;
	const dy = edit.endpoint === 'begin' ? oy - edit.y : edit.y - oy;
	const width = Math.hypot(dx, dy);
	if (!(width > 0))
		fail('UNSUPPORTED_GEOMETRY_EDIT', 'Coincident endpoint geometry is unverified.');
	const changed: VisioCellKey[] = [];
	for (const [axis, target] of [
		['X', edit.x],
		['Y', edit.y],
	] as const) {
		const name = `${prefix}${axis}`;
		if (numeric(local.get(name)) === target) continue;
		for (const lock of [`Lock${prefix}`, `LockMove${axis}`, 'LockWidth'])
			if (numeric(local.get(lock), 0) !== 0)
				fail('EDIT_PROTECTED_CELL', `${lock} prevents endpoint editing.`);
		editableCell(local.get(name));
		setCell(shape, name, target);
		changed.push({ pageId: edit.pageId, shapeId: edit.shapeId, cell: name });
	}
	return {
		changed,
		expected: { width, height: 0, x: (ox + edit.x) / 2, y: (oy + edit.y) / 2 },
		fixed: new Map([
			['Height', proof.get('Height')!],
			['Angle', Math.atan2(dy, dx)],
			['LocPinX', width / 2],
			['LocPinY', 0],
			['FlipX', 0],
			['FlipY', 0],
			[`${prefix}X`, edit.x],
			[`${prefix}Y`, edit.y],
			[`${opposite}X`, ox],
			[`${opposite}Y`, oy],
		] as const),
	};
}
