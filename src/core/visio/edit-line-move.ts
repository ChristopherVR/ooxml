import { attribute, children } from './sheet';
import { fail } from './package-common';
import { cells, editableCell, numeric, setCell } from './edit-geometry-admission';
import { executableCellFormula } from './cell-formula';
import type { VisioCellKey } from './edit-recalculate';
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
